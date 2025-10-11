# pip install edge-tts
# pip install moviepy


import os
import shutil
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import moviepy as mpy  # 保持原導入方式
import textwrap
from typing import List, Tuple
import re
import uuid
import math
import subprocess


# --- 內部輔助函數 ---

# ★★★ 新增的圖片處理函數 ★★★
import os
from PIL import Image


# ★★★ 修改後的圖片處理函數 ★★★
def _prepare_image(image_source: str | Image.Image, target_width: int, target_height: int) -> Image.Image:
    """
    將圖片處理後放置在指定尺寸的白色背景上。
    - 如果圖片小於目標尺寸，它將被置中，周圍是白色。
    - 如果圖片大於目標尺寸，它將按原比例縮小以適應目標尺寸，然後置中。

    Args:
        image_source (str | Image.Image): 圖片的路徑或已載入的 PIL Image 物件。
        target_width (int): 目標畫布寬度。
        target_height (int): 目標畫布高度。

    Returns:
        Image.Image: 處理後符合目標尺寸的 PIL Image 物件。
    """
    # 創建一個純白色的背景畫布
    background = Image.new("RGB", (target_width, target_height), color="white")

    img_to_paste = None
    try:
        # 根據傳入的是路徑還是已存在的 Image 物件來處理
        if isinstance(image_source, str):
            if not os.path.exists(image_source):
                raise FileNotFoundError
            img_to_paste = Image.open(image_source).convert("RGB")
        elif isinstance(image_source, Image.Image):
            img_to_paste = image_source.convert("RGB")

        if img_to_paste:
            # ★★★ 新增邏輯：如果圖片大於目標尺寸，則按比例縮小 ★★★
            original_width, original_height = img_to_paste.size
            if original_width > target_width or original_height > target_height:
                # 計算縮放比例，必須同時考慮寬度和高度，取較小的比例以確保縮放後能完整放入
                ratio = min(target_width / original_width, target_height / original_height)
                new_width = int(original_width * ratio)
                new_height = int(original_height * ratio)

                # 使用高品質的 LANCZOS 濾鏡進行縮放
                # (在舊版 Pillow 中可能使用 Image.ANTIALIAS)
                img_to_paste = img_to_paste.resize((new_width, new_height), Image.Resampling.LANCZOS)
            # ★★★ 結束新增邏輯 ★★★

            # 計算置中的位置 (無論是否縮放，都根據當前圖片尺寸來計算)
            paste_x = (target_width - img_to_paste.width) // 2
            paste_y = (target_height - img_to_paste.height) // 2

            # 將處理後的圖片粘貼到白色畫布上
            background.paste(img_to_paste, (paste_x, paste_y))
        else:
            print(f"警告: 無法處理提供的圖片來源，使用純白色背景。")

    except FileNotFoundError:
        print(f"警告: 圖片 '{image_source}' 未找到，使用純白色背景。")
    except Exception as e:
        print(f"警告: 載入或處理圖片 '{image_source}' 時出錯: {e}，使用純白色背景。")

    return background


def _wrap_text_and_adjust_font(text: str, draw: ImageDraw.Draw, font_path: str, initial_font_size: int, max_width: int,
                               max_height: int) -> tuple[str, ImageFont.FreeTypeFont]:
    # (此函數未變動)
    font_size = initial_font_size
    try:
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        print(f"警告: 字體 '{font_path}' 無法載入，將使用預設字體。")
        try:
            return text, ImageFont.load_default(size=font_size)
        except AttributeError:
            return text, ImageFont.load_default()
    while font_size > 10:
        try:
            avg_char_width = font.getlength("a")
        except AttributeError:
            # For older Pillow versions
            avg_char_width = font.getsize("a")[0]

        wrap_width = int(max_width / avg_char_width) if avg_char_width > 0 else 20
        wrapped_lines = textwrap.wrap(text, width=wrap_width, replace_whitespace=False)
        wrapped_text = "\n".join(wrapped_lines)
        bbox = draw.multiline_textbbox((0, 0), wrapped_text, font=font, spacing=4)
        text_height = bbox[3] - bbox[1]
        if text_height <= max_height:
            return wrapped_text, font
        font_size -= 2
        font = ImageFont.truetype(font_path, font_size)
    return wrapped_text, font


def _break_text_intelligently(text: str, font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    # (此函數未變動)
    final_lines = []
    text = text.replace('。', '. ').replace('！', '! ').replace('？', '? ').replace('，', ', ').replace('、', ', ')
    words = text.split()
    if not words:
        return []

    current_line = ""
    for word in words:
        test_line = (current_line + " " + word).strip()

        if font.getlength(test_line) <= max_width:
            current_line = test_line
        else:
            break_pos = -1
            for punc in ['.', '!', '?']:
                pos = current_line.rfind(punc)
                if pos > break_pos:
                    break_pos = pos

            if break_pos != -1 and break_pos > 0:
                line_to_add = current_line[:break_pos + 1]
                remaining_part = current_line[break_pos + 1:].strip()
                final_lines.append(line_to_add)
                current_line = (remaining_part + " " + word).strip()
            else:
                if current_line:
                    final_lines.append(current_line)
                current_line = word

    if current_line:
        final_lines.append(current_line)

    return [line.strip() for line in final_lines if line.strip()]


def _wrap_title_intelligently(text: str, num_lines: int) -> List[str]:
    # (此函數未變動)
    if num_lines <= 1:
        return [text]

    words = text.split()
    if not words or len(words) < num_lines:
        return text.splitlines() if '\n' in text else [text]

    lines = []
    word_idx = 0
    for i in range(num_lines):
        remaining_lines = num_lines - i
        remaining_words = len(words) - word_idx

        if remaining_lines == 1:
            words_for_this_line = remaining_words
        else:
            words_for_this_line = math.ceil(remaining_words / remaining_lines)
            words_for_this_line = min(words_for_this_line, remaining_words - (remaining_lines - 1))

        words_for_this_line = max(1, int(words_for_this_line))

        end_idx = word_idx + words_for_this_line
        lines.append(" ".join(words[word_idx:end_idx]))
        word_idx = end_idx

        if word_idx >= len(words):
            break

    return [line for line in lines if line]


def _generate_audio(text: str, lang: str, file_path: str) -> mpy.AudioFileClip | None:
    # (此函數未變動)
    if not text or not text.strip():
        print("警告: 嘗試為空文本生成音訊，已跳過。")
        return None

    voice_map = {
        'en': 'en-US-JennyNeural',
        'en-us': 'en-US-JennyNeural',
        'zh-tw': 'zh-TW-HsiaoChenNeural',
        'zh-cn': 'zh-CN-XiaoxiaoNeural',
        'ja-jp': 'ja-JP-NanamiNeural',
        'ko-kr': 'ko-KR-SunHiNeural',
    }
    voice = voice_map.get(lang.lower(), 'en-US-JennyNeural')

    command = [
        'edge-tts',
        '--voice', voice,
        '--text', text,
        '--write-media', file_path
    ]

    try:
        print(f"   -> 正在調用 Edge-TTS (CLI) 生成音訊...")
        subprocess.run(command, check=True, capture_output=True, text=True, encoding='utf-8')

        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            print(f"   -> ✅ Edge-TTS (CLI) 音訊生成成功: {os.path.basename(file_path)}")
            return mpy.AudioFileClip(file_path)
        else:
            print(f"   -> ❌ Edge-TTS (CLI) 錯誤: 生成的檔案 '{os.path.basename(file_path)}' 為空或不存在。")
            return None

    except FileNotFoundError:
        print("❌ 嚴重錯誤: 'edge-tts' 指令未找到。")
        print("   請確保您已運行 'pip install edge-tts' 並且 'edge-tts' 在系統的 PATH 環境變數中。")
        return None
    except subprocess.CalledProcessError as e:
        print(f"❌ 嚴重錯誤: 'edge-tts' 指令執行失敗 (返回碼: {e.returncode})。")
        print(f"   錯誤訊息: {e.stderr.strip()}")
        return None
    except Exception as e:
        print(f"❌ 在音訊生成期間發生未知錯誤: {e}")
        return None


def _generate_segmented_audio(lines: List[str], lang: str, temp_dir: str) -> Tuple[
    mpy.AudioFileClip | None, List[float]]:
    # (此函數未變動)
    if not lines:
        return None, []

    print(f"處理片段音訊 (共 {len(lines)} 行字幕)...")

    full_text = " ".join(lines)
    final_audio_path = os.path.join(temp_dir, f"FINAL_UNCUT_{uuid.uuid4().hex}.mp3")
    full_final_clip = _generate_audio(full_text, lang, final_audio_path)

    if full_final_clip is None:
        print(f"❌ 嚴重錯誤: 無法為片段生成完整音訊。")
        return None, []

    line_durations = []
    print("   -> 正在為每行字幕生成臨時音訊以精確計時...")
    for i, line in enumerate(lines):
        temp_line_path = os.path.join(temp_dir, f"TEMP_FOR_TIMING_{uuid.uuid4().hex}.mp3")
        temp_clip = None
        try:
            temp_clip = _generate_audio(line, lang, temp_line_path)

            if temp_clip is None:
                raise ValueError("無法生成臨時音訊檔。")

            duration = temp_clip.duration
            temp_clip.close()

            if not re.search(r'[.!?]\s*$', line):
                duration = max(duration - 0.6, 0.1)
            line_durations.append(duration)
        except Exception as e:
            print(f"⚠️ 警告: 第 {i} 行字幕的時長計算失敗: {e}，將使用 1.5 秒作為預設值。")
            line_durations.append(1.5)
        finally:
            if temp_clip:
                temp_clip.close()
            if os.path.exists(temp_line_path):
                os.remove(temp_line_path)

    line_end_times = np.cumsum(line_durations).tolist()
    total_calculated_duration = line_end_times[-1] if line_end_times else 0
    actual_full_duration = full_final_clip.duration

    if total_calculated_duration > 0 and abs(total_calculated_duration - actual_full_duration) > 0.1:
        print(
            f"   -> 總計算時長({total_calculated_duration:.2f}s)與實際音訊時長({actual_full_duration:.2f}s)存在差異，進行校準。")
        ratio = actual_full_duration / total_calculated_duration
        line_end_times = [t * ratio for t in line_end_times]

    if line_end_times:
        line_end_times[-1] = actual_full_duration

    return full_final_clip, line_end_times


# --- 畫面生成函數 ---

def create_typing_with_voiceover_clip(text: str, audio_clip: mpy.AudioFileClip, font_path: str, fps: int,
                                      pause_duration: float, font_size: int, video_width: int,
                                      video_height: int) -> mpy.VideoClip:
    # (此函數未變動)
    print("正在創建帶有同步配音的打字效果片段...")
    audio_duration = audio_clip.duration
    video_duration = audio_duration + pause_duration
    temp_draw = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    max_text_width = video_width - 100
    max_text_height = video_height - 100
    wrapped_text, final_font = _wrap_text_and_adjust_font(text, temp_draw, font_path, font_size, max_text_width,
                                                          max_text_height)
    clean_text = wrapped_text.replace('\n', '')
    total_chars = len(clean_text)
    char_per_second = total_chars / audio_duration if audio_duration > 0 else 0

    def make_frame(t):
        frame_img = Image.new("RGB", (video_width, video_height), color="black")
        draw = ImageDraw.Draw(frame_img)
        chars_to_show = int(t * char_per_second) if t < audio_duration else total_chars
        current_text_wrapped = ""
        count = 0
        for char in wrapped_text:
            if count >= chars_to_show: break
            current_text_wrapped += char
            if char != '\n': count += 1
        cursor = "|" if t < audio_duration and int(t * 2) % 2 == 0 else ""
        full_text_bbox = draw.multiline_textbbox((0, 0), wrapped_text, font=final_font, spacing=10, align="center")
        text_x = (video_width - (full_text_bbox[2] - full_text_bbox[0])) / 2
        text_y = (video_height - (full_text_bbox[3] - full_text_bbox[1])) / 2
        draw.multiline_text((text_x, text_y), current_text_wrapped + cursor, font=final_font, fill=(255, 255, 255),
                            align="center", spacing=10)
        return np.array(frame_img)

    return mpy.VideoClip(make_frame, duration=video_duration).with_audio(audio_clip)


# ★★★ 已修改的函數 ★★★
def create_styled_text_on_image_clip(
        image_path: str,
        text_to_display: str,
        voiceover_text: str,
        font_path: str,
        font_size: int,
        position: tuple,
        color: tuple,
        max_char_width: int,
        lang: str,
        temp_dir: str,
        video_width: int,
        video_height: int,
        pause_duration: float,
) -> mpy.VideoClip | None:
    """
    修改點：
    - 不再對圖片進行 .resize() 縮放操作。
    - 調用新的 _prepare_image 函數，將圖片置中放置在白色背景上。
    """
    print(f"正在創建風格化文字場景，文字: '{text_to_display}'")

    # 1. 生成音訊以確定片段時長
    audio_clip = _generate_audio(voiceover_text, lang, os.path.join(temp_dir, f"styled_{uuid.uuid4().hex}.mp3"))
    if not audio_clip:
        print(f"警告: 無法為 '{voiceover_text}' 生成配音，跳過此場景。")
        return None
    clip_duration = audio_clip.duration + pause_duration

    # 2. 創建視覺幀 (★★★ 修改點 ★★★)
    # 使用新的輔助函數來準備背景，不再進行縮放或拉伸
    background = _prepare_image(image_path, video_width, video_height)

    draw = ImageDraw.Draw(background)

    # 3. 繪製風格化文字
    try:
        font = ImageFont.truetype(font_path, font_size)
        wrapped_text = "\n".join(
            textwrap.wrap(text_to_display, width=max_char_width, replace_whitespace=False, break_long_words=True))
        draw.multiline_text(position, wrapped_text, font=font, fill=color, spacing=10, align="left")
    except Exception as e:
        print(f"警告: 無法繪製文字 '{text_to_display}': {e}")

    # 4. 創建並返回 moviepy 片段
    final_frame_rgb = background.convert("RGB")
    image_clip = mpy.ImageClip(np.array(final_frame_rgb)).with_duration(clip_duration)
    return image_clip.with_audio(audio_clip)


def create_clip_with_title_and_rolling_subtitles(
        base_image: Image.Image,
        video_title: str | None,
        full_text: str,
        lang: str,
        temp_dir: str,
        **kwargs
) -> mpy.VideoClip:
    # (此函數未變動，圖片處理已在其調用處完成)
    (title_font_path, subtitle_font_path, title_font_size, subtitle_font_size, font_color,
     bg_color, video_width, video_height, title_area_height, subtitle_area_height, pause_duration) = (
        kwargs['title_font_path'], kwargs['subtitle_font_path'], kwargs['title_font_size'],
        kwargs['subtitle_font_size'], kwargs['font_color'], kwargs['bg_color'],
        kwargs['video_width'], kwargs['video_height'], kwargs['title_area_height'],
        kwargs['subtitle_area_height'], kwargs['pause_duration']
    )

    try:
        title_font = ImageFont.truetype(title_font_path, title_font_size)
        subtitle_font = ImageFont.truetype(subtitle_font_path, subtitle_font_size)
    except IOError as e:
        print(f"錯誤: 無法加載字體: {e}")
        title_font = subtitle_font = ImageFont.load_default()

    # --- 1. 準備靜態背景 ---
    background_frame = Image.new("RGB", (video_width, video_height), "black")
    draw = ImageDraw.Draw(background_frame)
    img_y_start = title_area_height if video_title else 0
    subtitle_y_start = img_y_start + base_image.height

    if video_title:
        draw.rectangle([(0, 0), (video_width, img_y_start)], fill=(64, 64, 64))

    draw.rectangle([(0, subtitle_y_start), (video_width, video_height)], fill=bg_color)
    # 此處的 base_image 是已經由主函數處理好的，符合尺寸要求的圖片
    background_frame.paste(base_image, (0, img_y_start))

    # ★★★ 修改後的標題處理邏輯 ★★★
    if video_title:
        # 設定標題區域的最大寬度，左右各留 20px 的邊距
        title_max_width = video_width - 5
        current_font_size = title_font_size
        final_title_font = None

        # 循環調整字體大小，直到標題能在一行內放下
        while current_font_size > 10:  # 設定最小字體為 10
            try:
                final_title_font = ImageFont.truetype(title_font_path, current_font_size)
                # 獲取單行文本的總寬度
                text_width = final_title_font.getlength(video_title)

                # 如果寬度符合要求，則跳出循環
                if text_width <= title_max_width:
                    break
            except IOError:
                print(f"警告: 字體 '{title_font_path}' 在大小 {current_font_size} 時無法載入。")
                final_title_font = ImageFont.load_default()
                break  # 如果字體無法載入，則退出
            except Exception as e:
                print(f"錯誤: 獲取文本長度時出錯: {e}")
                break

            # 如果太寬，則減小字體大小並重試
            current_font_size -= 2

        if current_font_size <= 10 and final_title_font.getlength(video_title) > title_max_width:
            print(f"警告: 標題 '{video_title}' 在最小字體下仍然過寬。")

        # 計算繪製位置 (使其在灰色區域內水平和垂直居中)
        try:
            # Pillow >= 10.0.0, getbbox 更精確
            bbox = draw.textbbox((0, 0), video_title, font=final_title_font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
            # 計算垂直居中的y座標
            title_y = (title_area_height - text_h) / 2
        except AttributeError:
            # 舊版 Pillow
            text_w, text_h = draw.textsize(video_title, font=final_title_font)
            title_y = (title_area_height - text_h) / 2

        title_x = (video_width - text_w) / 2

        # 使用 draw.text 繪製單行文本
        draw.text((title_x, title_y), video_title, font=final_title_font, fill=font_color)
    # ★★★ 修改結束 ★★★

    # --- 2. 處理動態字幕和音訊 ---
    max_text_width = video_width - 40
    lines = _break_text_intelligently(full_text, subtitle_font, max_text_width)
    if not lines:
        return mpy.ImageClip(np.array(background_frame)).with_duration(kwargs.get("default_duration", 3.0))

    final_audio_clip, line_end_times = _generate_segmented_audio(lines, lang, temp_dir)
    if final_audio_clip is None:
        return mpy.ImageClip(np.array(background_frame)).with_duration(3.0)

    audio_duration = final_audio_clip.duration

    def make_frame(t):
        frame = background_frame.copy()
        draw = ImageDraw.Draw(frame)
        current_line_index = next((i for i, end_time in enumerate(line_end_times) if t < end_time), len(lines) - 1)
        current_line = lines[current_line_index]
        text_bbox = draw.textbbox((0, 0), current_line, font=subtitle_font)
        text_x = (video_width - (text_bbox[2] - text_bbox[0])) / 2
        text_y = subtitle_y_start + (subtitle_area_height - (text_bbox[3] - text_bbox[1])) / 2 - text_bbox[1]
        draw.text((text_x, text_y), current_line, font=subtitle_font, fill=font_color)
        return np.array(frame)

    return mpy.VideoClip(make_frame, duration=audio_duration + pause_duration).with_audio(final_audio_clip)


# ★★★ 已修改的函數 ★★★
def create_opening_scene_clips(
        image_path: str,
        video_title: str,
        thematic_opening_line: str,
        title_font_path: str,
        task_font_path: str,
        subtitle_font_path: str,
        subtitle_font_size: int,
        subtitle_font_color: tuple,
        subtitle_bg_color: tuple,
        subtitle_area_height: int,
        opening_title_font_size: int,
        opening_task_font_size: int,
        title_position: tuple,
        title_color: tuple,
        task_color: tuple,
        lang: str,
        temp_dir: str,
        video_width: int,
        video_height: int,
        pause_duration: float,
) -> List[mpy.VideoClip]:
    """
    創建開頭場景，包含兩個依序播放的片段。
    ★★★ 修改: 應用新的智能換行邏輯到 video_title。
    """
    clips = []
    max_text_width_for_title = (video_width // 4) * 3
    max_width_for_subtitle = video_width - 40

    # --- 1. 準備背景圖片和字體 ---
    try:
        base_bg = Image.open(image_path).resize((video_width, video_height), Image.Resampling.LANCZOS)
        font_title = ImageFont.truetype(title_font_path, opening_title_font_size)
        font_task = ImageFont.truetype(task_font_path, opening_task_font_size)
        font_subtitle = ImageFont.truetype(subtitle_font_path, subtitle_font_size)

        # ★★★ 修改後的標題換行邏輯 ★★★
        # 步驟 1: 估算理想行數
        try:
            avg_char_width_title = font_title.getlength("A") or opening_title_font_size / 2
        except AttributeError:
            avg_char_width_title = font_title.getsize("A")[0] or opening_title_font_size / 2

        wrap_width_title = max(1, int(max_text_width_for_title / avg_char_width_title))
        initial_lines = textwrap.wrap(video_title, width=wrap_width_title, replace_whitespace=False)
        num_lines = len(initial_lines)

        # 步驟 2: 使用新的智能換行函數來重新分配單詞
        final_lines = _wrap_title_intelligently(video_title, num_lines)
        wrapped_title = "\n".join(final_lines)
        # ★★★ 修改結束 ★★★

    except Exception as e:
        print(f"❌ 嚴重錯誤: 無法載入開頭場景字體或處理標題: {e}")
        return []

    # --- 2. 創建第一幕 (Thematic Line) ---
    if thematic_opening_line and video_title:
        print(f"正在創建開頭場景第一幕 (含字幕): '{video_title}'")

        lines_1 = _break_text_intelligently(thematic_opening_line, font_subtitle, max_width_for_subtitle)
        audio_clip_1, line_end_times_1 = _generate_segmented_audio(lines_1, lang, temp_dir)

        if audio_clip_1:
            static_frame_1 = base_bg.copy()
            draw1 = ImageDraw.Draw(static_frame_1)
            draw1.multiline_text(
                title_position, wrapped_title, font=font_title, fill=title_color, spacing=15, align="left"
            )

            def make_frame_1(t):
                frame = static_frame_1.copy().convert("RGBA")
                draw = ImageDraw.Draw(frame)
                sub_bg_y_start = video_height - subtitle_area_height
                draw.rectangle([(0, sub_bg_y_start), (video_width, video_height)], fill=subtitle_bg_color)
                current_line_index = next((i for i, end_time in enumerate(line_end_times_1) if t < end_time),
                                          len(lines_1) - 1)
                current_line = lines_1[current_line_index]
                text_bbox = draw.textbbox((0, 0), current_line, font=font_subtitle)
                text_x = (video_width - (text_bbox[2] - text_bbox[0])) / 2
                text_y = sub_bg_y_start + (subtitle_area_height - (text_bbox[3] - text_bbox[1])) / 2 - text_bbox[1]
                draw.text((text_x, text_y), current_line, font=font_subtitle, fill=subtitle_font_color)
                return np.array(frame.convert("RGB"))

            clip1 = mpy.VideoClip(make_frame_1, duration=audio_clip_1.duration + pause_duration).with_audio(
                audio_clip_1)
            clips.append(clip1)

    # # --- 3. 創建第二幕 (Sample Line) ---
    # if sample_opening_line and video_task:
    #     print(f"正在創建開頭場景第二幕 (含字幕): '{video_task}'")
    #
    #     lines_2 = _break_text_intelligently(sample_opening_line, font_subtitle, max_width_for_subtitle)
    #     audio_clip_2, line_end_times_2 = _generate_segmented_audio(lines_2, lang, temp_dir)
    #
    #     if audio_clip_2:
    #         static_frame_2 = base_bg.copy()
    #         draw2 = ImageDraw.Draw(static_frame_2)
    #         draw2.multiline_text(title_position, wrapped_title, font=font_title, fill=title_color, spacing=15,
    #                              align="left")
    #         title_bbox = draw2.multiline_textbbox(title_position, wrapped_title, font=font_title, spacing=15)
    #         task_x = title_position[0]
    #         task_y = title_bbox[3] + 20
    #         avg_char_width_task = font_task.getlength("A") or opening_task_font_size / 2
    #         wrap_width_task = max(1, int(max_text_width_for_title / avg_char_width_task))
    #         wrapped_task = "\n".join(textwrap.wrap(video_task, width=wrap_width_task, replace_whitespace=False))
    #         draw2.multiline_text((task_x, task_y), wrapped_task, font=font_task, fill=task_color, spacing=10,
    #                              align="left")
    #
    #         def make_frame_2(t):
    #             frame = static_frame_2.copy().convert("RGBA")
    #             draw = ImageDraw.Draw(frame)
    #             sub_bg_y_start = video_height - subtitle_area_height
    #             draw.rectangle([(0, sub_bg_y_start), (video_width, video_height)], fill=subtitle_bg_color)
    #             current_line_index = next((i for i, end_time in enumerate(line_end_times_2) if t < end_time),
    #                                       len(lines_2) - 1)
    #             current_line = lines_2[current_line_index]
    #             text_bbox = draw.textbbox((0, 0), current_line, font=font_subtitle)
    #             text_x = (video_width - (text_bbox[2] - text_bbox[0])) / 2
    #             text_y = sub_bg_y_start + (subtitle_area_height - (text_bbox[3] - text_bbox[1])) / 2 - text_bbox[1]
    #             draw.text((text_x, text_y), current_line, font=font_subtitle, fill=subtitle_font_color)
    #             return np.array(frame.convert("RGB"))
    #
    #         clip2 = mpy.VideoClip(make_frame_2, duration=audio_clip_2.duration + pause_duration).with_audio(
    #             audio_clip_2)
    #         clips.append(clip2)

    return clips


# ★★★ 主函數 (修改其內部圖片處理邏輯) ★★★
def create_video_with_subtitles_and_audio(
        video_title: str,
        thematic_opening_line: str,
        initial_image_path: str,
        image_text_map: dict,
        output_video_path: str,
        output_audio_folder: str,
        fps: int = 30,
        title_font_path: str = None,
        subtitle_font_path: str = None,
        title_font_size: int = 40,
        subtitle_font_size: int = 30,
        opening_title_font_size: int = 80,
        opening_task_font_size: int = 42,
        subtitle_font_color: tuple = (255, 255, 255),
        subtitle_bg_color: tuple = (0, 0, 0, 180),
        ending_sentence: str = None,
        lang: str = 'en'
):
    if not all([title_font_path, subtitle_font_path, initial_image_path]):
        raise FileNotFoundError("錯誤: 缺少必要的字體或初始圖片路徑。")

    # --- 統一定義尺寸 ---
    PAUSE_DURATION = 0.5
    TARGET_IMG_WIDTH, TARGET_IMG_HEIGHT = 1280, 720
    TITLE_AREA_HEIGHT, SUBTITLE_AREA_HEIGHT = 80, 100
    FINAL_VIDEO_WIDTH = TARGET_IMG_WIDTH
    CONTENT_VIDEO_HEIGHT = TITLE_AREA_HEIGHT + TARGET_IMG_HEIGHT + SUBTITLE_AREA_HEIGHT
    OPENING_SCENE_HEIGHT = TARGET_IMG_HEIGHT

    temp_dir = output_audio_folder
    os.makedirs(temp_dir, exist_ok=True)

    shared_kwargs = {
        "video_width": FINAL_VIDEO_WIDTH,
        "video_height": CONTENT_VIDEO_HEIGHT,
        "title_area_height": TITLE_AREA_HEIGHT,
        "subtitle_area_height": SUBTITLE_AREA_HEIGHT,
        "title_font_path": title_font_path,
        "subtitle_font_path": subtitle_font_path,
        "title_font_size": title_font_size,
        "subtitle_font_size": subtitle_font_size,
        "font_color": subtitle_font_color,
        "bg_color": subtitle_bg_color,
        "pause_duration": PAUSE_DURATION,
        "lang": lang,
        "temp_dir": temp_dir
    }

    try:
        video_clips = []

        # --- 1. 創建開頭場景 (已在其內部函數修改) ---
        opening_clips = create_opening_scene_clips(
            image_path=initial_image_path,
            video_title=video_title,
            thematic_opening_line=thematic_opening_line,
            title_font_path=title_font_path,
            task_font_path=subtitle_font_path,
            subtitle_font_path=subtitle_font_path,
            subtitle_font_size=subtitle_font_size,
            subtitle_font_color=subtitle_font_color,
            subtitle_bg_color=subtitle_bg_color,
            subtitle_area_height=SUBTITLE_AREA_HEIGHT,
            opening_title_font_size=opening_title_font_size,
            opening_task_font_size=opening_task_font_size,
            title_position=(100, 250),
            title_color=(31, 31, 31),
            task_color=(90, 90, 90),
            lang=lang,
            temp_dir=temp_dir,
            video_width=FINAL_VIDEO_WIDTH,
            video_height=OPENING_SCENE_HEIGHT,
            pause_duration=PAUSE_DURATION
        )
        if opening_clips:
            video_clips.extend(opening_clips)

        # --- 2. 處理中間的主要內容片段 ---
        for i, (image_key, data) in enumerate(image_text_map.items()):
            print(f"\n--- 處理片段 {i + 1}/{len(image_text_map)} ---")
            text = data.get("voiceover_script")
            title = data.get("title")
            if not text:
                continue

            # (★★★ 圖片處理邏輯修改 ★★★)
            base_image = None
            try:
                with Image.open(image_key) as img:
                    crop_pixels = 9
                    if img.width > crop_pixels * 2 and img.height > crop_pixels * 2:
                        left = crop_pixels
                        upper = crop_pixels
                        right = img.width - crop_pixels
                        lower = img.height - crop_pixels
                        print(f"   -> 裁切 '{os.path.basename(image_key)}' 四周各 {crop_pixels} 像素。")
                        cropped_img = img.crop((left, upper, right, lower))
                    else:
                        print(f"   -> 警告: 圖片 '{os.path.basename(image_key)}' 尺寸過小，無法裁切。")
                        cropped_img = img
                    # 使用新的輔助函數將裁切後的圖片置於白色畫布上，不再縮放
                    base_image = _prepare_image(cropped_img, TARGET_IMG_WIDTH, TARGET_IMG_HEIGHT)
            except FileNotFoundError:
                # 如果文件未找到，_prepare_image 會處理並返回一個白色背景
                print(f"警告: 圖片 '{image_key}' 未找到，使用白色背景替代。")
                base_image = _prepare_image(image_key, TARGET_IMG_WIDTH, TARGET_IMG_HEIGHT)
            # (★★★ 修改結束 ★★★)

            video_clip = create_clip_with_title_and_rolling_subtitles(
                base_image=base_image,
                video_title=title,
                full_text=text,
                **shared_kwargs
            )
            if video_clip:
                video_clips.append(video_clip)

        # --- 3. 創建結尾片段 ---
        if ending_sentence:
            ending_audio_clip = _generate_audio(ending_sentence, lang, os.path.join(temp_dir, "ending.mp3"))
            if ending_audio_clip:
                typing_clip = create_typing_with_voiceover_clip(
                    ending_sentence, ending_audio_clip, title_font_path, fps,
                    PAUSE_DURATION, title_font_size, FINAL_VIDEO_WIDTH, CONTENT_VIDEO_HEIGHT
                )
                video_clips.append(typing_clip)

        if not video_clips:
            print("❌ 錯誤: 沒有任何有效的片段可以合成影片。")
            return

        # --- 4. 合成最終影片 ---
        print("\n--- 正在合成最終視頻 ---")
        final_video = mpy.concatenate_videoclips(video_clips, method="compose")
        final_video.write_videofile(
            output_video_path, fps=fps, codec="libx264", audio_codec="aac", threads=4, logger='bar'
        )
        print(f"\n--- ✅ 視頻已成功創建並保存到: {output_video_path} ---")

    finally:
        pass


if __name__ == "__main__":
    # --- 字體和路徑設定 ---
    try:
        # Windows
        TITLE_FONT_PATH = 'C:/Windows/Fonts/msjhbd.ttc'  # 微軟正黑體 (粗體)
        SUBTITLE_FONT_PATH = 'C:/Windows/Fonts/msjh.ttc'  # 微軟正黑體 (常規)
        if not os.path.exists(TITLE_FONT_PATH):
            raise FileNotFoundError
    except FileNotFoundError:
        # macOS / Linux (備用)
        print("警告: 找不到 Windows 字體，嘗試使用 macOS/Linux 的備用字體。")
        TITLE_FONT_PATH = '/System/Library/Fonts/PingFang.ttc'
        SUBTITLE_FONT_PATH = '/System/Library/Fonts/PingFang.ttc'
        if not os.path.exists(TITLE_FONT_PATH):
            print("錯誤: 找不到任何可用的字體文件。請修改 `TITLE_FONT_PATH` 和 `SUBTITLE_FONT_PATH`。")
            exit()

    # --- 測試資料 ---
    # 1. 開頭場景的資料 (仿照範例圖片)
    video_title_text = "Microsoft Excel Microsoft Excel Excel"
    thematic_line = "Unlock more insightful data with our free, online, and collaborative spreadsheets."

    # 2. 初始背景圖路徑
    # !! 請務必將此路徑修改為您電腦上的圖片實際路徑 !!
    #    可以使用一張較小的圖片來測試白色背景效果。
    initial_image_file = r"C:\Users\v-yuhangxie\Desktop\excel2.jpg"
    initial_image_file_2 = r"C:\Users\v-yuhangxie\Desktop\excel3.jpg"

    # 3. 中間內容的資料 (使用相同圖片作為範例)
    image_data = {
        initial_image_file_2: {
            "voiceover_script": "Unlock more insightful data with our free, online, and collaborative spreadsheets.",
            "title": "Step 1: Introduction"
        },
        initial_image_file: {
            "voiceover_script": "This is the second part of our tutorial, where we explore advanced features.",
            "title": "Step 2: Advanced Features 111111111111111111111111111111122222222222222222222222222"
        }
    }
    # 4. 結尾句
    ending_text = "Thank you for watching!"

    output_file = "./final_video_no_scaling.mp4"
    output_audio_folder = "./audio_no_scaling"

    # --- 調用主函數 ---
    create_video_with_subtitles_and_audio(
        # 開頭場景參數
        video_title=video_title_text,
        thematic_opening_line=thematic_line,
        initial_image_path=initial_image_file,
        # 其他參數
        image_text_map=image_data,
        output_video_path=output_file,
        output_audio_folder=output_audio_folder,
        fps=24,
        title_font_path=TITLE_FONT_PATH,
        subtitle_font_path=SUBTITLE_FONT_PATH,
        # 這是後續內容的字體大小
        title_font_size=38,
        subtitle_font_size=32,
        # 這是為新開頭場景指定的字體大小
        opening_title_font_size=39,
        opening_task_font_size=27,
        lang='en',
        ending_sentence=ending_text
    )