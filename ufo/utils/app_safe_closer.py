# -*- coding: utf-8 -*-
import psutil
import os
import time
import logging

# Attempt to import Windows-specific libraries
_WINDOWS_AUTOMATION_ENABLED = False
_COM_ENABLED = False
try:
    import win32gui
    import win32process
    import win32con
    import pyautogui # Still needed for non-COM fallback / non-Office apps
    _WINDOWS_AUTOMATION_ENABLED = True
    print("Windows GUI automation libraries (pywin32, pyautogui) loaded.")
    try:
        import win32com.client
        import pywintypes # Import specifically for catching COM errors
        _COM_ENABLED = True
        print("Windows COM automation library (win32com) loaded.")
    except ImportError:
         logging.warning("Could not import win32com.client or pywintypes. COM-based Save As disabled.")

except ImportError:
    logging.warning("Could not import pywin32 or pyautogui. Graceful close/COM functionality will be disabled.")

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- COM Constants ---
# Word constants (WdSaveOptions)
WD_SAVE_OPTIONS = {
    "wdDoNotSaveChanges": 0,
    "wdSaveChanges": -1,
    "wdPromptToSaveChanges": -2,
}
# Excel constants (XlSaveAction)
XL_SAVE_ACTION = {
    "xlDoNotSaveChanges": 2,
    "xlSaveChanges": 1,
}
# PowerPoint constants (PpSaveAction)
PP_SAVE_ACTION = {
    "ppDoNotSaveChanges": 2,
    "ppSaveChanges": 1,
    "ppPromptToSaveChanges": 3,
}


class AppCloser:
    def __init__(self, close=True, object_name=None, app_name=None):
        """
        Initializes the AppCloser.
        Args:
            close (bool): Whether to actually attempt closing processes.
            object_name (str): Used to determine if targeting a specific app_name or standard_apps.
            app_name (str): The specific application name if object_name is set.
        """
        self.close = close
        self.object_name = object_name
        self.app_name = app_name
        self.standard_apps = ["WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"]
        # Timeout for GUI-based graceful close attempt (seconds)
        self.gui_graceful_close_timeout = 5
        # Map process names to their COM ProgIDs (Programmatic Identifiers)
        self.office_prog_ids = {
            "WINWORD.EXE": "Word.Application",
            "EXCEL.EXE": "Excel.Application",
            "POWERPNT.EXE": "PowerPoint.Application",
        }

    def _find_window_for_pid(self, pid):
        """Attempts to find a visible main window handle (HWND) for the given PID"""
        if not _WINDOWS_AUTOMATION_ENABLED:
            return None
        hwnd = None
        def enum_windows_callback(hWnd, lParam):
            nonlocal hwnd
            # Check if window is visible and a top-level window (more robust)
            if win32gui.IsWindowVisible(hWnd) and win32gui.GetParent(hWnd) == 0:
                 # Optional: Add more checks like checking class name if needed
                 _, found_pid = win32process.GetWindowThreadProcessId(hWnd)
                 if found_pid == pid:
                     # Ensure it has a title, often indicates a main window
                     if win32gui.GetWindowText(hWnd):
                         hwnd = hWnd
                         return False # Stop enumeration
            return True
        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            # Catch specific pywin32 errors if possible, e.g., pywintypes.error
            logging.error(f"Error enumerating windows for PID {pid}: {e}")
        return hwnd

    def _try_com_save_and_close(self, pid, process_name, save_path):
        """Attempts to Save As and Close using COM for Office apps."""
        if not _COM_ENABLED:
            logging.warning(f"COM SaveAs skipped for {process_name} (PID: {pid}): win32com not available.")
            return False

        prog_id = self.office_prog_ids.get(process_name.upper())
        if not prog_id:
            logging.warning(f"No COM ProgID known for {process_name}. Cannot use COM SaveAs.")
            return False

        print(f"Attempting COM SaveAs to '{save_path}' and Close for {process_name} (PID: {pid}) using ProgID '{prog_id}'...")

        app = None
        doc = None
        save_successful = False
        close_initiated = False # Track if close was attempted

        try:
            # --- Step 1: Connect to the COM object ---
            logging.info(f"Attempting COM connection via GetActiveObject('{prog_id}').")
            # Explicitly warn about the PID limitation
            logging.warning(f"Note: GetActiveObject may connect to *any* running instance, not necessarily PID {pid}.")
            try:
                app = win32com.client.GetActiveObject(prog_id)
                # app.Visible = True # Optionally make visible, might help responsiveness sometimes
                logging.info(f"Successfully connected to COM object for {prog_id}.")
            except pywintypes.com_error as com_err:
                logging.error(f"Failed to get active COM object for {prog_id}. COM Error: {com_err}", exc_info=True)
                return False # Connection failed, cannot proceed
            except Exception as e:
                 logging.error(f"An unexpected error occurred during GetActiveObject for {prog_id}: {e}", exc_info=True)
                 return False # Connection failed, cannot proceed

            # --- Check if connection was successful ---
            if app is None:
                # This case might be less likely if GetActiveObject raises an exception on failure, but check anyway
                logging.error("GetActiveObject returned None. Failed to connect to COM application.")
                return False

            # --- Step 2 & 3: Application Specific Logic with Enhanced Error Handling ---
            if process_name.upper() == "EXCEL.EXE":
                try:
                    # Check if Workbooks collection exists *before* accessing Count
                    if not hasattr(app, 'Workbooks'):
                         logging.error("Connected Excel COM object does not have 'Workbooks' attribute. Instance might be unusable.")
                         return False # Cannot proceed

                    workbooks_count = app.Workbooks.Count
                    print(f"Accessed app.Workbooks. Count: {workbooks_count}")

                    if workbooks_count == 0:
                        logging.warning("Excel application has COM object, but no workbooks are open.")
                        # Attempt to quit the empty application
                        try:
                            print("Attempting to Quit the empty Excel application via COM...")
                            app.Quit()
                            close_initiated = True # Mark quit as initiated
                            # Give it a bit more time to potentially close
                            time.sleep(4)
                            # Check if PID exists *after* quit attempt
                            if not psutil.pid_exists(pid):
                                print(f"Empty Excel process {pid} exited after COM Quit command.")
                                return True # Success, process is gone
                            else:
                                logging.warning(f"Excel process {pid} still exists after COM Quit for empty instance.")
                                return False # Quit command sent but didn't work, let fallback handle
                        except pywintypes.com_error as quit_com_err:
                            logging.warning(f"COM error sending Quit command to empty Excel: {quit_com_err}")
                            return False
                        except Exception as quit_e:
                            logging.warning(f"Sending COM Quit command failed: {quit_e}")
                            return False

                    # Try to get the ActiveWorkbook
                    if not hasattr(app, 'ActiveWorkbook'):
                         logging.error("Connected Excel COM object does not have 'ActiveWorkbook' attribute.")
                         # Maybe iterate app.Workbooks(1) ? More complex. For now, fail.
                         return False

                    doc = app.ActiveWorkbook
                    if doc is None:
                        logging.warning("Excel has open workbooks, but no ActiveWorkbook identified via COM (ActiveWorkbook is None). Cannot determine which to save.")
                        return False # Unclear which workbook to target

                    doc_name = "N/A" # Default name in case of error getting it
                    try:
                         doc_name = doc.Name
                         print(f"Identified Active Workbook via COM: {doc_name}")

                         # Perform SaveAs
                         print(f"Attempting SaveAs for '{doc_name}' to '{save_path}'...")
                         doc.SaveAs(save_path)
                         save_successful = True
                         print(f"Workbook '{doc_name}' saved via COM to '{save_path}'.")

                         # Perform Close
                         print(f"Attempting Close for Workbook '{doc_name}'...")
                         doc.Close(SaveChanges=False) # Don't save changes again after SaveAs
                         close_initiated = True
                         print(f"Workbook '{doc_name}' closed via COM.")
                         doc = None # Release doc object

                         # Check if app should be quit (last workbook closed?)
                         try:
                             if app.Workbooks.Count == 0:
                                 print("Last workbook closed. Attempting to Quit Excel application.")
                                 app.Quit()
                                 time.sleep(5) # Give time for Quit to process
                                 # Check if PID exists *after* quit attempt
                                 if not psutil.pid_exists(pid):
                                     print(f"Excel process {pid} exited after COM Quit command.")
                                     app = None
                                     return True # Quit succeeded, process gone
                                 else:
                                     logging.warning(f"Excel process {pid} still exists after final COM Quit command.")
                                     # Return True because the main goal (SaveAs/Close doc) succeeded
                                     # Fallback might still be needed for the process itself
                                     return True
                             else:
                                 print(f"Other workbooks ({app.Workbooks.Count}) still open. Not quitting Excel application.")
                                 # Return True as the target workbook's operation succeeded
                                 return True
                         except pywintypes.com_error as check_quit_com_err:
                              logging.warning(f"COM error checking workbook count or quitting app after closing doc: {check_quit_com_err}")
                              return True # Return True as the main operation succeeded
                         except Exception as check_quit_e:
                              logging.warning(f"Error checking workbook count or quitting app after closing doc: {check_quit_e}")
                              return True # Return True as the main operation succeeded

                    except pywintypes.com_error as save_close_com_err:
                         logging.error(f"COM Error during SaveAs or Close for Workbook '{doc_name}': {save_close_com_err}", exc_info=True)
                         return False # COM operation failed
                    except Exception as save_close_e:
                         logging.error(f"Error during COM SaveAs or Close for Workbook '{doc_name}': {save_close_e}", exc_info=True)
                         return False # Other error during operation

                except AttributeError as ae:
                    # Catch attribute errors specifically during Excel property access
                    logging.error(f"AttributeError accessing Excel COM properties (e.g., 'Workbooks', 'ActiveWorkbook'): {ae}. COM object might be invalid or Excel state issue.", exc_info=True)
                    return False
                except pywintypes.com_error as excel_com_err:
                     # Catch COM errors during general Excel operations
                     logging.error(f"A COM error occurred during Excel operations: {excel_com_err}", exc_info=True)
                     return False
                except Exception as e:
                     # Catch any other unexpected errors
                     logging.error(f"An unexpected error occurred during Excel COM handling: {e}", exc_info=True)
                     return False


            # --- Add similar robust blocks for Word and PowerPoint ---
            elif process_name.upper() == "WINWORD.EXE":
                try:
                    # Check Documents collection
                    if not hasattr(app, 'Documents') or not hasattr(app.Documents, 'Count'):
                        logging.error("Connected Word COM object missing 'Documents' attribute or 'Count'.")
                        return False
                    if app.Documents.Count == 0:
                        logging.warning("Word application has COM object, but no documents open.")
                        # Optional: Add Quit logic for empty Word instance here if desired
                        return False # Or True if empty means "nothing to save"

                    # Check ActiveDocument
                    if not hasattr(app, 'ActiveDocument'):
                        logging.error("Connected Word COM object missing 'ActiveDocument' attribute.")
                        return False
                    doc = app.ActiveDocument
                    if doc is None:
                         logging.warning("Word has documents open, but ActiveDocument is None.")
                         return False

                    doc_name = getattr(doc, 'Name', 'N/A')
                    print(f"Identified Active Document via COM: {doc_name}")

                    try:
                        # SaveAs
                        print(f"Attempting SaveAs for '{doc_name}' to '{save_path}'...")
                        doc.SaveAs(save_path)
                        save_successful = True
                        print(f"Document '{doc_name}' saved via COM to '{save_path}'.")

                        # Close
                        print(f"Attempting Close for Document '{doc_name}'...")
                        doc.Close(SaveChanges=WD_SAVE_OPTIONS["wdDoNotSaveChanges"])
                        close_initiated = True
                        print(f"Document '{doc_name}' closed via COM.")
                        doc = None
                        # Optional: Add conditional app.Quit() logic here similar to Excel
                        print("Word document closed. Checking if app should quit (logic not implemented here).")
                        return True # Assume success means doc handled

                    except pywintypes.com_error as save_close_com_err:
                         logging.error(f"COM Error during SaveAs or Close for Document '{doc_name}': {save_close_com_err}", exc_info=True)
                         return False
                    except Exception as save_close_e:
                         logging.error(f"Error during COM SaveAs or Close for Document '{doc_name}': {save_close_e}", exc_info=True)
                         return False

                except AttributeError as ae:
                     logging.error(f"AttributeError accessing Word COM properties: {ae}.", exc_info=True)
                     return False
                except pywintypes.com_error as word_com_err:
                     logging.error(f"A COM error occurred during Word operations: {word_com_err}", exc_info=True)
                     return False
                except Exception as word_e:
                     logging.error(f"Unexpected error during Word COM handling: {word_e}", exc_info=True)
                     return False

            elif process_name.upper() == "POWERPNT.EXE":
                 # Similar structure for PowerPoint
                try:
                    if not hasattr(app, 'Presentations') or not hasattr(app.Presentations, 'Count'):
                        logging.error("Connected PowerPoint COM object missing 'Presentations' or 'Count'.")
                        return False
                    if app.Presentations.Count == 0:
                        logging.warning("PowerPoint application has COM object, but no presentations open.")
                        # Optional Quit logic
                        return False

                    if not hasattr(app, 'ActivePresentation'):
                         logging.error("Connected PowerPoint COM object missing 'ActivePresentation'.")
                         return False
                    doc = app.ActivePresentation
                    if doc is None:
                         logging.warning("PowerPoint has presentations open, but ActivePresentation is None.")
                         return False

                    doc_name = getattr(doc, 'Name', 'N/A')
                    print(f"Identified Active Presentation via COM: {doc_name}")

                    try:
                        # SaveAs
                        print(f"Attempting SaveAs for '{doc_name}' to '{save_path}'...")
                        # Note: PowerPoint SaveAs might require explicit format type sometimes
                        # Example: doc.SaveAs(save_path, FileFormat=1) # ppSaveAsDefault
                        doc.SaveAs(save_path)
                        save_successful = True
                        print(f"Presentation '{doc_name}' saved via COM to '{save_path}'.")

                        # Close
                        print(f"Attempting Close for Presentation '{doc_name}'...")
                        doc.Close()
                        close_initiated = True
                        print(f"Presentation '{doc_name}' closed via COM.")
                        doc = None
                        # Optional conditional Quit
                        print("PowerPoint presentation closed. Checking if app should quit (logic not implemented here).")
                        return True

                    except pywintypes.com_error as save_close_com_err:
                         logging.error(f"COM Error during SaveAs or Close for Presentation '{doc_name}': {save_close_com_err}", exc_info=True)
                         return False
                    except Exception as save_close_e:
                         logging.error(f"Error during COM SaveAs or Close for Presentation '{doc_name}': {save_close_e}", exc_info=True)
                         return False

                except AttributeError as ae:
                     logging.error(f"AttributeError accessing PowerPoint COM properties: {ae}.", exc_info=True)
                     return False
                except pywintypes.com_error as ppt_com_err:
                     logging.error(f"A COM error occurred during PowerPoint operations: {ppt_com_err}", exc_info=True)
                     return False
                except Exception as ppt_e:
                     logging.error(f"Unexpected error during PowerPoint COM handling: {ppt_e}", exc_info=True)
                     return False

            else:
                logging.warning(f"COM logic not implemented for {process_name}")
                return False

        except Exception as outer_e:
            # Catch-all for any other unexpected errors during the process
            logging.error(f"Unhandled exception during COM attempt for PID {pid} ({process_name}): {outer_e}", exc_info=True)
            return False # COM attempt failed overall

        finally:
            # --- Release COM objects ---
            # Explicitly release COM objects to be safe, especially in error cases
            # Python's garbage collector should handle this, but explicit is clearer
            # Note: Releasing 'app' might prevent Quit() in some scenarios if not handled carefully above.
            # The logic above now returns True/False, so this finally block mainly cleans up dangling references if an exception occurred mid-way.
            if doc is not None:
                try:
                    # No standard release method, rely on garbage collection after setting to None
                    doc = None
                    logging.debug(f"Set doc object to None for PID {pid}")
                except Exception as rel_e:
                    logging.warning(f"Exception while trying to release doc object (set to None): {rel_e}")
            if app is not None:
                try:
                    # Only set app to None; avoid calling Quit() here as it's handled specifically above
                    app = None
                    logging.debug(f"Set app object to None for PID {pid}")
                except Exception as rel_e:
                     logging.warning(f"Exception while trying to release app object (set to None): {rel_e}")

            # The function should have returned True or False within the try block based on success/failure.
            # If execution reaches here unexpectedly (e.g., missing return path), assume failure.
            logging.debug(f"COM function finally block reached for PID {pid}.")
            # Return based on whether save and close were at least attempted/successful
            # This return is a fallback; ideally, returns happen within the try block.
            # If we initiated a close/quit, we consider the attempt made.
            # return close_initiated or save_successful # Or just return False if reaching here is an error state


    def _try_graceful_close_windows_gui(self, pid, process_name):
        """Attempts to gracefully close a Windows application using GUI simulation (Ctrl+S, Alt+F4)."""
        if not _WINDOWS_AUTOMATION_ENABLED:
            logging.warning(f"GUI Graceful close skipped for {process_name} (PID: {pid}): Automation libraries not available.")
            return False

        print(f"Attempting GUI graceful close for {process_name} (PID: {pid})...")
        hwnd = self._find_window_for_pid(pid)

        if not hwnd:
            logging.warning(f"Could not find window for GUI graceful close (PID: {pid}).")
            return False

        try:
            print(f"Activating window {hwnd} for GUI close...")
            try:
                 # Try different methods to bring window to front
                 # win32gui.ShowWindow(hwnd, win32con.SW_RESTORE) # Can sometimes be disruptive
                 win32gui.SetForegroundWindow(hwnd)
                 time.sleep(0.5) # Short pause after activation
                 # Optional: Use pyautogui activation as fallback
                 # pyautogui.click(win32gui.GetWindowRect(hwnd)[0], win32gui.GetWindowRect(hwnd)[1])
            except pywintypes.error as fg_err:
                 # Specifically catch error if window is already in foreground or cannot be activated
                 logging.warning(f"SetForegroundWindow failed (Error code: {fg_err.winerror}). Trying to continue GUI attempt... Error: {fg_err}")
                 # Fallback using pyautogui activate (less reliable)
                 try:
                      pyautogui.getWindowsWithTitle(win32gui.GetWindowText(hwnd))[0].activate()
                      time.sleep(0.5)
                 except Exception as pa_activate_e:
                      logging.warning(f"PyAutoGUI activate also failed: {pa_activate_e}")
            except Exception as e:
                 logging.warning(f"Activating window failed for HWND {hwnd}: {e}. Continuing GUI attempt...")

            pyautogui.PAUSE = 0.2 # Pause between pyautogui actions

            # Send Save (Ctrl+S) - This might open a "Save As" dialog if file is new
            print(f"Sending GUI 'Save' (Ctrl+S) to PID {pid} (HWND: {hwnd})")
            pyautogui.hotkey('ctrl', 's')
            time.sleep(1.5) # Give time for save dialog or save operation

            # Send Close (Alt+F4) - This should close the window or prompt if unsaved
            print(f"Sending GUI 'Close' (Alt+F4) to PID {pid} (HWND: {hwnd})")
            pyautogui.hotkey('alt', 'f4')
            time.sleep(0.5) # Allow close message to process

            # Check for standard save prompts (Optional, makes it more complex)
            # You could add pyautogui image recognition here to handle "Save", "Don't Save", "Cancel" prompts
            # print("Checking for save prompts (logic not implemented)...")

            print(f"Waiting up to {self.gui_graceful_close_timeout} seconds for PID {pid} to close (GUI attempt)...")
            # Check more frequently initially
            for i in range(self.gui_graceful_close_timeout * 2): # Check every 0.5s
                if not psutil.pid_exists(pid):
                    print(f"Process {pid} ({process_name}) closed after GUI graceful attempt.")
                    return True
                time.sleep(0.5)

            logging.warning(f"Process {pid} ({process_name}) did not close within timeout after GUI graceful attempt.")
            return False
        except Exception as e:
            logging.error(f"Exception during GUI graceful close attempt for PID {pid} ({process_name}): {e}", exc_info=True)
            return False

    def terminate_application_processes(self, save_path_map=None):
        """
        Terminates specific application processes.
        For Office apps, attempts COM SaveAs if save_path_map provides a path for the PID.
        Falls back to GUI graceful close (Ctrl+S, Alt+F4).
        Finally uses force kill if needed.
        """
        if not self.close:
            print("Process termination skipped (self.close is False).")
            return

        target_app_names = []
        if self.object_name and self.app_name:
            # Target a specific application if provided
            target_app_names = [self.app_name.upper()]
            print(f"Targeting specific application: {self.app_name}")
        else:
            # Target the default list of standard applications
            target_app_names = [name.upper() for name in self.standard_apps]
            print(f"Targeting standard applications: {', '.join(self.standard_apps)}")

        processes_to_terminate = []
        try:
            # Iterate through running processes to find targets
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    pinfo = proc.info
                    # Check if process info is valid and name matches target list
                    if pinfo and pinfo.get('name') and pinfo['name'].upper() in target_app_names:
                        processes_to_terminate.append(pinfo)
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    # Ignore processes that can't be accessed or are gone
                    continue
                except Exception as proc_iter_e: # Catch errors getting info for a single process
                     logging.warning(f"Could not get info for a process during iteration: {proc_iter_e}")
                     continue
        except Exception as e:
             logging.error(f"Error iterating processes: {e}")
             return # Cannot proceed if process iteration fails

        if not processes_to_terminate:
            print("No target application processes found running.")
            return

        print(f"Found {len(processes_to_terminate)} target process(es) to terminate.")

        if save_path_map is None:
            save_path_map = {} # Initialize if not provided

        # Process each found target application instance
        for pinfo in processes_to_terminate:
            pid = pinfo.get('pid')
            process_name = pinfo.get('name')

            # Defensive check for valid PID/Name obtained from iteration
            if not pid or not process_name:
                 logging.warning(f"Skipping invalid process info entry: {pinfo}")
                 continue

            # Check if the process still exists before attempting to close
            if not psutil.pid_exists(pid):
                print(f"Process {pid} ({process_name}) already closed before processing.")
                continue

            print(f"\n--- Processing PID: {pid} ({process_name}) ---")
            graceful_attempt_made = False # Track if COM or GUI was tried
            process_closed = False # Track if the process is confirmed closed

            # --- Attempt 1: COM Save As (if path provided and Office app) ---
            target_save_path = save_path_map.get(pid)
            if target_save_path and process_name.upper() in self.office_prog_ids:
                 if os.name == 'nt' and _COM_ENABLED:
                     print(f"Attempting COM SaveAs/Close for PID {pid}...")
                     com_success = self._try_com_save_and_close(pid, process_name, target_save_path)
                     graceful_attempt_made = True
                     if com_success:
                         print(f"COM SaveAs/Close attempt for PID {pid} reported success.")
                         # Verify if process actually closed after COM attempt
                         time.sleep(1) # Short delay for process exit
                         if not psutil.pid_exists(pid):
                             print(f"Process {pid} confirmed closed after successful COM operation.")
                             process_closed = True
                         else:
                              # COM might succeed for the document but leave the process running
                              logging.warning(f"Process {pid} still running after COM operation reported success. May need fallback.")
                     else:
                          print(f"COM SaveAs/Close attempt failed or was not applicable for PID {pid}.")
                 else:
                     # Log why COM was skipped
                     logging.warning(f"COM SaveAs for PID {pid} skipped (Not Windows, COM disabled, or no path).")

            # --- Attempt 2: GUI Graceful Close (if COM didn't close the process or wasn't applicable) ---
            if not process_closed:
                # Check if the process still exists before trying GUI
                if psutil.pid_exists(pid):
                    print(f"Attempting GUI graceful close for {process_name} (PID: {pid})...")
                    if os.name == 'nt' and _WINDOWS_AUTOMATION_ENABLED:
                        gui_success = self._try_graceful_close_windows_gui(pid, process_name)
                        graceful_attempt_made = True
                        if gui_success:
                            # _try_graceful_close_windows_gui already confirms pid_exists is false on success
                             print(f"GUI graceful close successful for PID {pid}.")
                             process_closed = True # GUI function confirmed it closed
                        else:
                             print(f"GUI graceful close failed or timed out for PID {pid}.")
                    else:
                         # Log why GUI was skipped
                         print(f"Skipping GUI graceful close for {process_name} (PID: {pid}) (Not Windows or GUI libs missing).")
                else:
                    # Process closed between COM and GUI attempts (maybe COM Quit worked delayed)
                    print(f"Process {pid} ({process_name}) closed before GUI fallback needed.")
                    process_closed = True

            # --- Attempt 3: Force Kill (if all graceful attempts failed/skipped and process still exists) ---
            if not process_closed:
                 if psutil.pid_exists(pid): # Check again before killing
                    # Only force kill if a graceful attempt was actually made and failed,
                    # or if no graceful method was available/applicable.
                    if graceful_attempt_made or not (os.name == 'nt' and (_COM_ENABLED or _WINDOWS_AUTOMATION_ENABLED)):
                         logging.warning(f"Graceful close failed or skipped for {process_name} (PID: {pid}). Attempting force kill...")
                         try:
                             proc_to_kill = psutil.Process(pid)
                             proc_to_kill.kill() # Send kill signal
                             print(f"Sent kill signal to process {pid}.")
                             time.sleep(2) # Wait for kill signal to be processed
                             if not psutil.pid_exists(pid):
                                  print(f"Process {pid} ({process_name}) successfully force-killed.")
                                  process_closed = True
                             else:
                                  # Kill might fail for various reasons (permissions, zombie)
                                  logging.error(f"Force kill command seemed to fail for PID {pid}. Process still exists.")
                         except psutil.NoSuchProcess:
                              # Process terminated between pid_exists check and kill()
                              print(f"Process {pid} was already gone before force kill executed.")
                              process_closed = True
                         except Exception as e:
                              logging.error(f"Exception during force kill for PID {pid} ({process_name}): {e}")
                    else:
                         # Avoid force kill if graceful methods weren't even attempted (likely config issue)
                         logging.warning(f"Skipping force kill for PID {pid} as no graceful attempt was made (likely config issue).")

                 else:
                    # Process closed between GUI and Kill attempts
                    print(f"Process {pid} ({process_name}) closed before force kill was needed.")
                    process_closed = True

            # Final status logging for this PID
            if not process_closed and psutil.pid_exists(pid):
                 logging.error(f"Failed to close PID {pid} ({process_name}) after all attempts.")
            elif process_closed:
                 print(f"Successfully processed PID {pid} ({process_name}).")
            else:
                 # This case means process_closed is False but pid doesn't exist anymore (closed during checks)
                 print(f"Process PID {pid} ({process_name}) appears closed after processing cycle.")


# ==============================================================================
# Simplified Interface Function (Relies on the improved class methods)
# ==============================================================================
def save_and_close_app(app_name: str, save_as_path: str) -> bool:
    """
    Finds the first running instance of the specified application,
    attempts to save its active document/workbook/presentation to the
    specified path using COM (for Office apps), then attempts to close it
    gracefully (including quitting the app if it becomes empty for Excel),
    falling back to GUI close and finally force kill if necessary.

    Args:
        app_name (str): The process name (e.g., "EXCEL.EXE", "WINWORD.EXE").
        save_as_path (str): The full, absolute path to save the file to.
                           The directory must exist.

    Returns:
        bool: True if the target process is no longer running after the
              attempts, False otherwise (or if the process wasn't found).
              Note: True doesn't guarantee the save operation itself was
              successful if fallback methods were used, only that the process
              is likely closed. Check logs for details.
    """
    # --- Input Validation ---
    if not app_name or not save_as_path:
        logging.error("Application name and save path must be provided.")
        return False

    # --- Path Validation ---
    if not os.path.isabs(save_as_path):
         # Absolute paths are generally safer for COM SaveAs
         logging.warning(f"Save path '{save_as_path}' is not absolute. This might cause issues.")
    save_dir = os.path.dirname(save_as_path)
    if not os.path.isdir(save_dir):
        logging.error(f"Save directory '{save_dir}' does not exist.")
        # Optionally create the directory:
        # try:
        #     os.makedirs(save_dir)
        #     print(f"Created save directory: {save_dir}")
        # except OSError as e:
        #     logging.error(f"Failed to create save directory '{save_dir}': {e}")
        #     return False
        return False # Return False if directory doesn't exist (and not created)

    # --- Find Target Process ---
    target_pid = None
    target_pname = None
    try:
        app_name_upper = app_name.upper()
        # Iterate to find the *first* matching process instance
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                pinfo = proc.info
                # Robust check for name existence before comparing
                if pinfo and pinfo.get('name') and pinfo['name'].upper() == app_name_upper:
                    target_pid = pinfo['pid']
                    target_pname = pinfo['name'] # Store name for logging
                    print(f"Found target process {target_pname} with PID: {target_pid}")
                    break # Target the first instance found
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue # Ignore processes we can't access
            except Exception as proc_find_e:
                 logging.warning(f"Error getting info for process during search: {proc_find_e}")
                 continue

        if target_pid is None:
            # If no process is running, the goal (app closed) is already met
            print(f"No running process found for application name: {app_name}")
            return True
    except Exception as e:
        logging.error(f"Error finding process for {app_name}: {e}", exc_info=True)
        return False

    # --- Prepare and Execute Closure Attempt ---
    # Use object_name to indicate it's targeting a specific app via this interface
    closer = AppCloser(close=True, object_name="SingleAppInterface", app_name=app_name)
    # Map the found PID to the desired save path
    save_map = {target_pid: save_as_path}

    print(f"Calling termination logic for PID {target_pid} ({target_pname}) with save path '{save_as_path}'")
    final_process_exists = True
    try:
        # Call the core method that handles the COM/GUI/Kill sequence
        closer.terminate_application_processes(save_path_map=save_map)

        # --- Final Check and Return Value ---
        # Check the final state of the specific target process after all attempts
        # Add a slightly longer delay to allow force kill etc. to register fully
        time.sleep(3)
        final_process_exists = psutil.pid_exists(target_pid)

        if not final_process_exists:
            print(f"Process PID {target_pid} successfully confirmed closed.")
            return True # Success: The target process is gone
        else:
            logging.warning(f"Process PID {target_pid} still running after all attempts in terminate_application_processes.")
            return False # Failure: The target process persists
    except Exception as e:
         # Catch errors in this function's calling logic itself
         logging.error(f"An error occurred within the save_and_close_app function structure: {e}", exc_info=True)
         # Check PID status even if an error occurred somewhere in the call stack
         if target_pid:
             final_process_exists = psutil.pid_exists(target_pid)
             if not final_process_exists:
                 # Process closed despite error elsewhere
                 print(f"Process PID {target_pid} seems closed despite error during termination call.")
                 return True # Return True as the process is gone
             else:
                 logging.warning(f"Process PID {target_pid} likely still running after error in termination call.")
                 return False
         else:
              # If target_pid wasn't even found initially due to error
              return False


# ==============================================================================
# Example Usage (Adjust paths as needed)
# ==============================================================================
if __name__ == "__main__":

    print("\n--- Example: Using save_and_close_app() for Excel ---")
    # !!! IMPORTANT: Ensure Excel is running with a workbook open !!!
    # !!! IMPORTANT: Use an absolute path and ensure the DIRECTORY exists or change path!!!
    target_app_excel = "EXCEL.EXE"
    timestamp = int(time.time())
    # Make sure C:\temp exists or change this path to a valid one on your system
    excel_save_path = rf"C:\temp\ClosedViaInterface_Excel_{timestamp}.xlsx"

    print(f"Attempting to save and close {target_app_excel}...")
    print(f"Target save path: {excel_save_path}")

    # Ensure target directory exists before calling the function
    excel_save_dir = os.path.dirname(excel_save_path)
    if not os.path.exists(excel_save_dir):
         print(f"Creating directory: {excel_save_dir}")
         os.makedirs(excel_save_dir, exist_ok=True)

    # Call the simplified interface function
    was_closed_excel = save_and_close_app(target_app_excel, excel_save_path)

    # Report outcome
    if was_closed_excel:
        print(f"Successfully closed or verified closure of {target_app_excel} process (PID might have changed if multiple instances existed). Check logs for save details.")
    else:
        print(f"Failed to close {target_app_excel} process. Check logs for errors.")

    print("------------------------------------------------------\n")

    # --- Optional: Example for Word ---
    # print("\n--- Example: Using save_and_close_app() for Word ---")
    # target_app_word = "WINWORD.EXE"
    # # Use a valid path on your system
    # target_path_word = rf"C:\temp\ClosedViaInterface_Word_{int(time.time())}.docx"
    # print(f"Attempting to save and close {target_app_word}...")
    # print(f"Target save path: {target_path_word}")
    # target_dir_word = os.path.dirname(target_path_word)
    # if not os.path.exists(target_dir_word):
    #     print(f"Creating directory: {target_dir_word}")
    #     os.makedirs(target_dir_word, exist_ok=True)
    # was_closed_word = save_and_close_app(target_app_word, target_path_word)
    # if was_closed_word:
    #     print(f"Successfully closed or verified closure of {target_app_word} process(es).")
    # else:
    #     print(f"Failed to close {target_app_word} process.")
    # print("------------------------------------------------------\n")