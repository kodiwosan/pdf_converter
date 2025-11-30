import os
import subprocess
import time

import cv2
import img2pdf
import numpy as np
import pyautogui
import pygetwindow as gw
import pytesseract
from PIL import Image, ImageChops
from pypdf import PdfReader, PdfWriter


def list_windows():
    """タイトルがあるウィンドウのリストを返します。"""
    return [w for w in gw.getAllWindows() if w.title]


class KindleConverter:
    def __init__(self, window_title="Kindle"):
        self.window_title = window_title
        self.window = None
        self.crop_region = None  # (left, top, width, height)
        self.temp_dir = "temp_screenshots"
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)
        self.setup_tesseract()

    def setup_tesseract(self):
        """Tesseractの実行ファイルを一般的なパスから探して設定します。"""
        if pytesseract.pytesseract.tesseract_cmd != "tesseract":
            pass
        else:
            common_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                os.path.expanduser(r"~\AppData\Local\Tesseract-OCR\tesseract.exe"),
                os.path.expanduser(
                    r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
                ),
            ]

            for path in common_paths:
                if os.path.exists(path):
                    print(f"Tesseract found at: {path}")
                    pytesseract.pytesseract.tesseract_cmd = path
                    break
            else:
                print("Warning: Tesseract executable not found in common paths.")
                print(
                    "Please ensure Tesseract-OCR is installed and added to PATH, or installed in a standard location."
                )
                return

        # 言語データの確認
        try:
            langs = pytesseract.get_languages()
            print(f"Available Tesseract languages: {langs}")
        except Exception as e:
            print(f"Warning: Could not get Tesseract languages: {e}")

    def locate_window(self):
        """Kindleのウィンドウを探してアクティブにします。"""
        windows = gw.getWindowsWithTitle(self.window_title)
        if not windows:
            raise Exception(f"Window with title '{self.window_title}' not found.")

        self.window = windows[0]
        try:
            self.window.activate()
        except Exception as e:
            print(f"Warning: Could not activate window directly: {e}")
            if self.window.isMinimized:
                self.window.restore()

        time.sleep(1)
        print(f"Located window: {self.window.title}")

    def auto_detect_crop_region(self):
        """
        コンテンツ領域を自動検出します。
        2値化（Otsu）を行い、検出された輪郭のうち「内部が白（ページ色）で満たされているもの」を抽出します。
        これにより、黒い帯を含むウィンドウ枠や、小さなUIパーツを除外します。
        """
        print("Attempting auto-detection of content region (White Density Method)...")

        # ウィンドウ全体のスクリーンショットを取得
        region = (
            self.window.left,
            self.window.top,
            self.window.width,
            self.window.height,
        )
        screenshot = pyautogui.screenshot(region=region)
        img_np = np.array(screenshot)
        img_h, img_w = img_np.shape[:2]

        # OpenCV形式に変換
        img_cv = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

        # 2値化 (大津の二値化)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        _, thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # 輪郭抽出
        contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

        page_contours = []

        # UIの高さの目安
        ui_top_limit = img_h * 0.10
        ui_bottom_limit = img_h * 0.90

        debug_img = img_cv.copy()

        for cnt in contours:
            x, y, w, h = cv2.boundingRect(cnt)
            area = w * h

            # 1. 小さすぎる領域は除外 (ウィンドウの5%未満)
            if area < (img_w * img_h * 0.05):
                continue

            # 2. UI領域（上下端）にあるものは除外
            center_y = y + h / 2
            if center_y < ui_top_limit or center_y > ui_bottom_limit:
                cv2.rectangle(
                    debug_img, (x, y), (x + w, y + h), (255, 0, 0), 2
                )  # 青: UI除外
                continue

            # 3. 「白の密度」を計算
            roi = thresh[y : y + h, x : x + w]
            white_pixels = cv2.countNonZero(roi)
            density = white_pixels / area

            # ページであれば、大部分が白（テキストがあっても70%以上は白）のはず
            if density < 0.7:
                print(
                    f"Skipping contour at ({x},{y}) due to low white density: {density:.2f}"
                )
                cv2.rectangle(
                    debug_img, (x, y), (x + w, y + h), (0, 255, 255), 2
                )  # 黄: 密度不足
                continue

            # 合格したものをページ候補とする
            page_contours.append((x, y, w, h))
            cv2.rectangle(debug_img, (x, y), (x + w, y + h), (0, 255, 0), 2)  # 緑: 採用

        if not page_contours:
            print("No suitable page contours found. Using full window.")
            cv2.imwrite(
                os.path.join(self.temp_dir, "auto_crop_debug_density.png"), debug_img
            )
            return

        # 採用された全輪郭を包含する矩形を求める (Union)
        min_x = min([c[0] for c in page_contours])
        min_y = min([c[1] for c in page_contours])
        max_x = max([c[0] + c[2] for c in page_contours])
        max_y = max([c[1] + c[3] for c in page_contours])

        # パディング
        padding = 5
        min_x = max(0, min_x - padding)
        min_y = max(0, min_y - padding)
        max_x = min(img_w, max_x + padding)
        max_y = min(img_h, max_y + padding)

        crop_w = max_x - min_x
        crop_h = max_y - min_y

        # ウィンドウ相対座標から絶対座標へ変換
        abs_x = self.window.left + min_x
        abs_y = self.window.top + min_y

        self.crop_region = (abs_x, abs_y, crop_w, crop_h)
        print(f"Auto-detected crop region: {self.crop_region}")

        # 確認用
        test_shot = pyautogui.screenshot(region=self.crop_region)
        test_path = os.path.join(self.temp_dir, "auto_crop_test.png")
        test_shot.save(test_path)

        # デバッグ画像
        cv2.rectangle(
            debug_img, (min_x, min_y), (max_x, max_y), (0, 0, 255), 3
        )  # 赤: 最終結果
        cv2.imwrite(
            os.path.join(self.temp_dir, "auto_crop_result_debug.png"), debug_img
        )

        print(f"Saved auto-crop test to {test_path}.")

    def calibrate_crop_region(self):
        """ユーザーに対話的にクロッピング領域を指定させます。"""
        print("\n--- Crop Calibration ---")
        print("1. Move your mouse to the TOP-LEFT corner of the book content.")
        input("Press Enter to capture Top-Left position...")
        x1, y1 = pyautogui.position()
        print(f"Top-Left captured: ({x1}, {y1})")

        print("2. Move your mouse to the BOTTOM-RIGHT corner of the book content.")
        input("Press Enter to capture Bottom-Right position...")
        x2, y2 = pyautogui.position()
        print(f"Bottom-Right captured: ({x2}, {y2})")

        left = min(x1, x2)
        top = min(y1, y2)
        width = abs(x2 - x1)
        height = abs(y2 - y1)

        self.crop_region = (left, top, width, height)
        print(f"Crop region set: {self.crop_region}")

        test_shot = pyautogui.screenshot(region=self.crop_region)
        test_path = os.path.join(self.temp_dir, "test_crop.png")
        test_shot.save(test_path)
        print(f"Saved test screenshot to {test_path}. Please check it.")
        input("Press Enter to continue if the crop is correct (or Ctrl+C to abort)...")

    def get_window_region(self):
        if self.crop_region:
            return self.crop_region
        if not self.window:
            raise Exception("Window not located. Call locate_window() first.")
        return (
            self.window.left,
            self.window.top,
            self.window.width,
            self.window.height,
        )

    def capture_page(self, page_num):
        region = self.get_window_region()
        screenshot_path = os.path.join(self.temp_dir, f"page_{page_num:04d}.png")
        screenshot = pyautogui.screenshot(region=region)
        screenshot.save(screenshot_path)
        print(f"Captured page {page_num}")
        return screenshot_path

    def next_page(self):
        pyautogui.press("right")
        time.sleep(1.5)

    def is_same_image(self, img1_path, img2_path):
        if not img1_path or not img2_path:
            return False
        try:
            img1 = Image.open(img1_path)
            img2 = Image.open(img2_path)
            if img1.size != img2.size:
                return False
            diff = ImageChops.difference(img1, img2)
            if not diff.getbbox():
                return True
            return False
        except Exception as e:
            print(f"Error comparing images: {e}")
            return False

    def convert_to_pdf(self, output_filename="output.pdf", use_ocr=False, lang="jpn"):
        print("Converting to PDF...")
        image_files = sorted(
            [
                os.path.join(self.temp_dir, f)
                for f in os.listdir(self.temp_dir)
                if f.endswith(".png")
            ]
        )

        if not image_files:
            print("No images found to convert.")
            return

        if use_ocr:
            print(f"Performing OCR (Language: {lang}). This may take a while...")

            # 言語がインストールされているか確認
            try:
                available_langs = pytesseract.get_languages()
                if lang not in available_langs:
                    print(f"Error: Language '{lang}' is not installed in Tesseract.")
                    print(f"Available languages: {available_langs}")
                    print("Falling back to standard image-only PDF.")
                    with open(output_filename, "wb") as f:
                        f.write(img2pdf.convert(image_files))
                    return
            except:
                pass  # 無視して続行

            merger = PdfWriter()
            temp_pdfs = []

            # Tesseractコマンドの取得
            tesseract_cmd = pytesseract.pytesseract.tesseract_cmd

            try:
                for i, img_path in enumerate(image_files):
                    base_name = os.path.splitext(img_path)[0]
                    # Tesseract CLI出力ベース名（拡張子なし）
                    output_base = base_name

                    # 期待されるPDFパス
                    pdf_path = output_base + ".pdf"

                    # CLIコマンド: tesseract <image> <output_base> -l <lang> pdf
                    cmd = [tesseract_cmd, img_path, output_base, "-l", lang, "pdf"]

                    try:
                        # CLI実行
                        subprocess.run(
                            cmd,
                            check=True,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                        )

                        # ファイル生成確認
                        if not os.path.exists(pdf_path):
                            # Tesseractのバージョンによっては.pdfがつかない、あるいは別の名前になる可能性も考慮
                            # しかし通常は output_base.pdf になる
                            print(f"Warning: Expected PDF not found at {pdf_path}")
                            continue

                        # 生成されたPDFにテキストが含まれているか検証 (全ページ)
                        char_count = 0
                        try:
                            reader = PdfReader(pdf_path)
                            if len(reader.pages) > 0:
                                text = reader.pages[0].extract_text()
                                char_count = len(text)
                        except Exception as val_e:
                            print(
                                f"Warning: Could not verify PDF text for {os.path.basename(img_path)}: {val_e}"
                            )

                        print(
                            f"OCR processing ({i + 1}/{len(image_files)}): {os.path.basename(img_path)} - {char_count} chars extracted"
                        )

                        merger.append(pdf_path)
                        temp_pdfs.append(pdf_path)

                    except subprocess.CalledProcessError as e:
                        print(
                            f"Tesseract CLI Error for {os.path.basename(img_path)}: {e}"
                        )
                        continue

                merger.write(output_filename)
                merger.close()
                print(f"Searchable PDF saved to {output_filename}")

                # 最終PDFの検証
                try:
                    print("Verifying final PDF content...")
                    reader = PdfReader(output_filename)
                    total_chars = 0
                    for page in reader.pages:
                        total_chars += len(page.extract_text())
                    print(
                        f"Final PDF Verification: {len(reader.pages)} pages, {total_chars} total characters."
                    )
                except Exception as e:
                    print(f"Could not verify final PDF: {e}")

                # Clean up temp pdfs (Debug: Disabled to allow inspection)
                print(
                    "Debug: Intermediate PDFs kept in temp_screenshots for inspection."
                )
                # for p in temp_pdfs:
                #     if os.path.exists(p):
                #         os.remove(p)

            except Exception as e:
                print(f"OCR Error: {e}")
                import traceback

                traceback.print_exc()
                print("Falling back to standard image-only PDF.")
                with open(output_filename, "wb") as f:
                    f.write(img2pdf.convert(image_files))
        else:
            with open(output_filename, "wb") as f:
                f.write(img2pdf.convert(image_files))
            print(f"PDF saved to {output_filename}")

    def run(self, use_ocr=False, crop_mode="none"):
        try:
            self.locate_window()

            if crop_mode == "manual":
                self.calibrate_crop_region()
                self.locate_window()
            elif crop_mode == "auto":
                self.auto_detect_crop_region()
                self.locate_window()

            print("Starting in 3 seconds... Please ensure Kindle window is visible.")
            time.sleep(3)

            page_num = 1
            last_screenshot_path = None

            while True:
                current_screenshot_path = self.capture_page(page_num)

                if last_screenshot_path and self.is_same_image(
                    last_screenshot_path, current_screenshot_path
                ):
                    print("End of book detected.")
                    os.remove(current_screenshot_path)
                    break

                last_screenshot_path = current_screenshot_path
                self.next_page()
                page_num += 1

                if page_num > 2000:
                    print("Safety limit reached.")
                    break

            self.convert_to_pdf(use_ocr=use_ocr)
            print("Done!")

        except Exception as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    pass
