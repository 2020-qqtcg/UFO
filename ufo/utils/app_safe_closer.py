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
        _COM_ENABLED = True
        print("Windows COM automation library (win32com) loaded.")
    except ImportError:
         logging.warning("Could not import win32com.client. COM-based Save As disabled.")

except ImportError:
    logging.warning("Could not import pywin32 or pyautogui. Graceful close/COM functionality will be disabled.")

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Constants for COM ---
# Word constants (WdSaveOptions) can be looked up if needed, wdDoNotSaveChanges = 0, wdSaveChanges = -1, wdPromptToSaveChanges = -2
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
        self.close = close
        self.object_name = object_name # Used to decide if specific app_name or standard_apps is targeted
        self.app_name = app_name # Specific app if object_name is set
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
        """Attempt to find a visible main window handle (HWND) for the given PID"""
        # (Same implementation as before)
        if not _WINDOWS_AUTOMATION_ENABLED:
            return None
        hwnd = None
        def enum_windows_callback(hWnd, lParam):
            nonlocal hwnd
            if win32gui.IsWindowVisible(hWnd) and win32gui.IsWindowEnabled(hWnd):
                _, found_pid = win32process.GetWindowThreadProcessId(hWnd)
                if found_pid == pid:
                    if win32gui.GetWindowText(hWnd):
                        hwnd = hWnd
                        return False
            return True
        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
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

        print(f"Attempting COM SaveAs to '{save_path}' and Close for {process_name} (PID: {pid})...")

        app = None
        doc = None
        try:
            # --- Step 1: Connect to the COM object ---
            logging.warning(f"Attempting to connect via COM using GetActiveObject('{prog_id}'). "
                            f"This might not connect to the specific instance with PID {pid} if multiple instances are running.")
            try:
                # app = win32com.client.GetActiveObject(prog_id)
                app = win32com.client.Dispatch(prog_id)
                if app is None:
                    logging.error("GetActiveObject returned None. Failed to connect to COM application.")
                    return False
            except Exception as com_connect_e:
                 logging.error(f"Failed to get active COM object for {prog_id}: {com_connect_e}", exc_info=True)
                 return False

            # --- Step 2: Access Application Specific Properties ---
            if process_name.upper() == "EXCEL.EXE":
                try:
                    workbooks_collection = app.Workbooks
                    print(f"Successfully accessed app.Workbooks. Count: {workbooks_collection.Count}")
                    if workbooks_collection.Count == 0:
                         logging.warning("Excel application has COM object, but no workbooks are open.")
                         try:
                             print("Attempting to Quit the empty Excel application via COM...")
                             app.Quit()
                             time.sleep(3)
                             if not psutil.pid_exists(pid):
                                 print(f"Empty Excel process {pid} exited after COM Quit command.")
                                 return True
                             else:
                                 logging.warning(f"Excel process {pid} still exists after COM Quit command.")
                                 return False # Quit didn't terminate, let fallback handle it
                         except Exception as quit_e:
                             logging.warning(f"Sending COM Quit command failed: {quit_e}")
                             return False
                    doc = app.ActiveWorkbook
                    if doc is None:
                        logging.warning("Excel has open workbooks, but no ActiveWorkbook identified via COM. Cannot determine which to save.")
                        return False
                except AttributeError as ae:
                    logging.error(f"AttributeError accessing COM properties (likely 'Workbooks'): {ae}. COM object might be invalid or Excel in bad state.", exc_info=True)
                    app = None
                    return False
                except Exception as e:
                     logging.error(f"An unexpected error occurred when accessing Excel workbooks/properties via COM: {e}", exc_info=True)
                     app = None
                     return False

                # --- Step 3: Perform SaveAs and Close on the Document/Workbook ---
                try:
                    doc_name = doc.Name
                    print(f"Identified Active Workbook via COM: {doc_name}")
                    doc.SaveAs(save_path)
                    print(f"Workbook '{doc_name}' saved via COM to '{save_path}'.")
                    doc.Close(SaveChanges=False)
                    print(f"Workbook '{doc_name}' closed via COM.")
                    doc = None # Release doc object

                    # **** MODIFICATION START ****
                    # Check if closing this workbook left the app empty and quit if so
                    try:
                        if app.Workbooks.Count == 0:
                            print("Last workbook closed. Attempting to Quit Excel application.")
                            app.Quit()
                            time.sleep(5) # Give time for Quit to process
                            if not psutil.pid_exists(pid):
                                print(f"Excel process {pid} exited after COM Quit command.")
                                app = None # Release app object
                                return True # Quit succeeded, process gone
                            else:
                                logging.warning(f"Excel process {pid} still exists after COM Quit command was sent.")
                                app = None
                                # Even if Quit didn't kill process, main goal achieved
                                return True
                        else:
                            # If other workbooks remain, don't quit the application
                            print(f"Other workbooks ({app.Workbooks.Count}) still open in this Excel instance. Not quitting application.")
                            app = None # Release app object
                            # Return True as the target workbook's operation succeeded
                            return True
                    except Exception as check_quit_e:
                         logging.warning(f"Error checking workbook count or quitting app after closing doc: {check_quit_e}")
                         app = None # Release app object
                         # Return True as the target workbook's operation succeeded
                         return True
                    # **** MODIFICATION END ****

                except Exception as save_close_e:
                     logging.error(f"Error during COM SaveAs or Close for Workbook '{getattr(doc, 'Name', 'N/A')}': {save_close_e}", exc_info=True)
                     doc = None
                     app = None
                     return False # SaveAs or Close failed

            # --- Add similar blocks for Word and PowerPoint ---
            # (You might want to apply similar conditional Quit logic for Word/PowerPoint too)
            elif process_name.upper() == "WINWORD.EXE":
                 # Simplified logic - Add refinement as needed
                 try:
                     if not app.Documents.Count > 0:
                         logging.warning("Word application has COM object, but no documents are open.")
                         # Optionally add app.Quit() logic here too
                         return False
                     doc = app.ActiveDocument
                     doc_name = doc.Name # Get name before potential close
                     print(f"Identified Active Document via COM: {doc_name}")
                     doc.SaveAs(save_path)
                     print(f"Document '{doc_name}' saved via COM to '{save_path}'.")
                     doc.Close(SaveChanges=WD_SAVE_OPTIONS["wdDoNotSaveChanges"])
                     print(f"Document '{doc_name}' closed via COM.")
                     doc = None
                     # Add conditional app.Quit() logic here similar to Excel if desired
                     print("Word document closed. Checking if app should quit (logic not implemented here).")
                     # For now, assume success means doc handled, app state is secondary
                     app = None
                     return True
                 except Exception as word_e:
                     logging.error(f"Error during Word COM operation: {word_e}", exc_info=True)
                     doc = None
                     app = None
                     return False


            elif process_name.upper() == "POWERPNT.EXE":
                 # Simplified logic - Add refinement as needed
                 try:
                     if not app.Presentations.Count > 0:
                         logging.warning("PowerPoint application has COM object, but no presentations are open.")
                         # Optionally add app.Quit() logic here too
                         return False
                     doc = app.ActivePresentation
                     doc_name = doc.Name # Get name before potential close
                     print(f"Identified Active Presentation via COM: {doc_name}")
                     doc.SaveAs(save_path)
                     print(f"Presentation '{doc_name}' saved via COM to '{save_path}'.")
                     doc.Close() # PowerPoint Close is on Presentation, no SaveChanges arg needed after SaveAs
                     print(f"Presentation '{doc_name}' closed via COM.")
                     doc = None
                     # Add conditional app.Quit() logic here similar to Excel if desired
                     print("PowerPoint presentation closed. Checking if app should quit (logic not implemented here).")
                     # For now, assume success means doc handled, app state is secondary
                     app = None
                     return True
                 except Exception as ppt_e:
                     logging.error(f"Error during PowerPoint COM operation: {ppt_e}", exc_info=True)
                     doc = None
                     app = None
                     return False

            else:
                logging.warning(f"COM logic not implemented for {process_name}")
                # Explicitly release app object if obtained
                app = None
                return False

            # Note: The code execution should ideally return from within the if/elif blocks now.
            # This final section might only be reached if logic is added above that doesn't return.

            # --- Step 4: Final Check (Potentially redundant now due to returns in blocks above) ---
            # This part is less likely to be reached for Office apps if the logic above returns True/False correctly.
            # time.sleep(1) # Give time for any async operations
            # if not psutil.pid_exists(pid):
            #    print(f"Process {pid} ({process_name}) exited after COM operations.")
            #    app = None # Ensure app object is released
            #    return True
            # else:
            #    print(f"Process {pid} ({process_name}) still running after COM operations.")
            #    app = None # Release app object
            #    return True # Return True as the primary COM action (SaveAs/Close) was likely attempted/succeeded


        except Exception as outer_e:
            # Catch any unexpected errors in the overall process
            logging.error(f"Unhandled exception during COM attempt for PID {pid} ({process_name}): {outer_e}", exc_info=True)
            doc = None
            app = None
            return False # COM attempt failed

    def _try_graceful_close_windows_gui(self, pid, process_name):
        """Attempts to gracefully close a Windows application using GUI simulation (Ctrl+S, Alt+F4)."""
        # (Same implementation as before)
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
                 win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                 win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                 logging.warning(f"SetForegroundWindow failed for HWND {hwnd}: {e}. Continuing GUI attempt...")
                 pyautogui.PAUSE = 0.2
            time.sleep(2)
            print(f"Sending GUI 'Save' (Ctrl+S) to PID {pid}")
            pyautogui.hotkey('ctrl', 's')
            time.sleep(2)
            print(f"Sending GUI 'Close' (Alt+F4) to PID {pid}")
            pyautogui.hotkey('alt', 'f4')
            print(f"Waiting up to {self.gui_graceful_close_timeout} seconds for PID {pid} to close (GUI attempt)...")
            for _ in range(self.gui_graceful_close_timeout):
                time.sleep(2)
                if not psutil.pid_exists(pid):
                    print(f"Process {pid} ({process_name}) closed after GUI graceful attempt.")
                    return True
            logging.warning(f"Process {pid} ({process_name}) did not close after GUI graceful attempt.")
            return False
        except Exception as e:
            logging.error(f"Exception during GUI graceful close attempt for PID {pid} ({process_name}): {e}")
            return False

    def terminate_application_processes(self, save_path_map=None):
        """
        Terminates specific application processes.
        For Office apps, attempts COM SaveAs if save_path_map provides a path for the PID.
        Falls back to GUI graceful close (Ctrl+S, Alt+F4).
        Finally uses force kill if needed.
        """
        # (Implementation remains the same as before)
        if not self.close:
            print("Process termination skipped (self.close is False).")
            return

        target_app_names = []
        if self.object_name and self.app_name:
            target_app_names = [self.app_name.upper()]
            print(f"Targeting specific application: {self.app_name}")
        else:
            target_app_names = [name.upper() for name in self.standard_apps]
            print(f"Targeting standard applications: {', '.join(self.standard_apps)}")

        processes_to_terminate = []
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    pinfo = proc.info
                    if pinfo['name'].upper() in target_app_names:
                        processes_to_terminate.append(pinfo)
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
        except Exception as e:
             logging.error(f"Error iterating processes: {e}")
             return

        if not processes_to_terminate:
            print("No target application processes found running.")
            return

        print(f"Found {len(processes_to_terminate)} target process(es) to terminate.")

        if save_path_map is None:
            save_path_map = {}

        for pinfo in processes_to_terminate:
            pid = pinfo['pid']
            process_name = pinfo['name']

            if not psutil.pid_exists(pid):
                print(f"Process {pid} ({process_name}) already closed before processing.")
                continue

            graceful_success = False

            # --- Attempt 1: COM Save As (if path provided and Office app) ---
            target_save_path = save_path_map.get(pid)
            if target_save_path and process_name.upper() in self.office_prog_ids:
                 if os.name == 'nt' and _COM_ENABLED:
                     graceful_success = self._try_com_save_and_close(pid, process_name, target_save_path)
                 else:
                     logging.warning(f"COM SaveAs for PID {pid} skipped (Not Windows or COM disabled).")

            # --- Attempt 2: GUI Graceful Close (if COM didn't succeed or wasn't applicable) ---
            if not graceful_success:
                # Check if the process still exists before trying GUI, COM might have failed but closed it
                if psutil.pid_exists(pid):
                    print(f"Falling back to GUI graceful close attempt for {process_name} (PID: {pid}).")
                    if os.name == 'nt' and _WINDOWS_AUTOMATION_ENABLED:
                        graceful_success = self._try_graceful_close_windows_gui(pid, process_name)
                    else:
                         print(f"Skipping GUI graceful close for {process_name} (PID: {pid}) (Not Windows or GUI libs missing).")
                else:
                    print(f"Process {pid} ({process_name}) already closed before GUI fallback needed.")
                    # We consider this a success in terms of the process being closed, even if COM failed initially.
                    graceful_success = True # Set to true so force kill isn't attempted

            # --- Attempt 3: Force Kill (if all else failed and process still exists) ---
            if not graceful_success:
                if psutil.pid_exists(pid): # Check again before killing
                    logging.warning(f"All graceful close attempts failed or skipped for {process_name} (PID: {pid}). Attempting force kill...")
                    try:
                        if os.name == 'nt':
                            kill_command = f"taskkill /f /pid {pid}"
                            print(f"Executing force kill: {kill_command}")
                            result = os.system(kill_command)
                            time.sleep(2)
                            if psutil.pid_exists(pid):
                                 logging.error(f"Force kill command seemed to fail for PID {pid} (Return code: {result}).")
                            else:
                                 print(f"Process {pid} ({process_name}) successfully force-killed.")
                        else:
                            logging.warning(f"Force kill not implemented for OS '{os.name}' for PID {pid}")
                    except Exception as e:
                        logging.error(f"Exception during force kill for PID {pid} ({process_name}): {e}")
                else:
                    print(f"Process {pid} ({process_name}) closed before force kill was needed.")

# ==============================================================================
# NEW Simplified Interface Function
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
    if not app_name or not save_as_path:
        logging.error("Application name and save path must be provided.")
        return False

    # Basic path validation
    if not os.path.isabs(save_as_path):
         logging.warning(f"Save path '{save_as_path}' is not absolute. This might cause issues.")
    save_dir = os.path.dirname(save_as_path)
    if not os.path.isdir(save_dir):
        logging.error(f"Save directory '{save_dir}' does not exist.")
        # Optionally, attempt to create it:
        # try:
        #     os.makedirs(save_dir)
        #     print(f"Created save directory: {save_dir}")
        # except OSError as e:
        #     logging.error(f"Failed to create save directory '{save_dir}': {e}")
        #     return False
        return False # Return False if directory doesn't exist

    target_pid = None
    try:
        app_name_upper = app_name.upper()
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'].upper() == app_name_upper:
                    target_pid = proc.info['pid']
                    print(f"Found target process {app_name} with PID: {target_pid}")
                    break # Target the first instance found
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        if target_pid is None:
            print(f"No running process found for application name: {app_name}")
            return True # No process running, so considered "closed"
    except Exception as e:
        logging.error(f"Error finding process for {app_name}: {e}", exc_info=True)
        return False

    # Prepare for the AppCloser call
    # Use object_name to indicate it's targeting a specific app
    closer = AppCloser(close=True, object_name="SingleAppInterface", app_name=app_name)
    save_map = {target_pid: save_as_path}

    print(f"Calling internal termination logic for PID {target_pid} with save path '{save_as_path}'")
    try:
        # Call the method that handles the COM/GUI/Kill sequence
        closer.terminate_application_processes(save_path_map=save_map)

        # Check the final state of the process
        # Add a small delay to allow final termination attempts to register
        time.sleep(5)
        if not psutil.pid_exists(target_pid):
            print(f"Process PID {target_pid} successfully closed.")
            return True
        else:
            logging.warning(f"Process PID {target_pid} still running after all attempts.")
            return False
    except Exception as e:
         logging.error(f"An error occurred during the termination process call: {e}", exc_info=True)
         # Check PID status even if error occurred during the call
         if target_pid and not psutil.pid_exists(target_pid):
             print(f"Process PID {target_pid} seems closed despite error during termination call.")
             return True
         else:
             logging.warning(f"Process PID {target_pid} likely still running after error in termination call.")
             return False


# ==============================================================================
# Example Usage of the New Interface
# ==============================================================================
if __name__ == "__main__":

    print("\n--- Example: Using save_and_close_app() for Excel ---")
    # Ensure Excel is running with a workbook open

    # Define the target application and save path
    target_app = "EXCEL.EXE"
    # !!! IMPORTANT: Use an absolute path and ensure C:\temp exists or change path!!!
    timestamp = int(time.time())
    target_path = rf"D:\ClosedViaInterface_{timestamp}.xlsx"

    print(f"Attempting to save and close {target_app}...")
    print(f"Target save path: {target_path}")

    # Ensure directory exists (or handle potential error from function)
    target_dir = os.path.dirname(target_path)
    if not os.path.exists(target_dir):
         print(f"Creating directory: {target_dir}")
         os.makedirs(target_dir, exist_ok=True)

    # Call the interface function
    was_closed = save_and_close_app(target_app, target_path)

    if was_closed:
        print(f"Successfully closed or verified closure of {target_app} process(es). Check logs for save details.")
    else:
        print(f"Failed to close {target_app} process. Check logs for errors.")

    print("------------------------------------------------------\n")

    # --- Optional: Example for Word ---
    # print("\n--- Example: Using save_and_close_app() for Word ---")
    # target_app_word = "WINWORD.EXE"
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