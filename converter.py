import os
import time

import img2pdf
import pyautogui
import pygetwindow as gw
from PIL import Image, ImageChops


def list_windows():
    """タイトルがあるウィンドウのリストを返します。"""
    return [w for w in gw.getAllWindows() if w.title]


class KindleConverter:
    def __init__(self, window_title="Kindle"):
        self.window_title = window_title
        self.window = None
        self.temp_dir = "temp_screenshots"
        if not os.path.exists(self.temp_dir):
            os.makedirs(self.temp_dir)

    def locate_window(self):
        """Kindleのウィンドウを探してアクティブにします。"""
        windows = gw.getWindowsWithTitle(self.window_title)
        if not windows:
            raise Exception(f"Window with title '{self.window_title}' not found.")

        # 複数の候補がある場合は最初のものを使用
        self.window = windows[0]

        # ウィンドウをアクティブにする
        try:
            self.window.activate()
        except Exception as e:
            print(f"Warning: Could not activate window directly: {e}")
            if self.window.isMinimized:
                self.window.restore()

        time.sleep(1)  # アクティブになるのを待つ
        print(f"Located window: {self.window.title}")

    def get_window_region(self):
        """ウィンドウの位置とサイズを取得します。"""
        if not self.window:
            raise Exception("Window not located. Call locate_window() first.")

        return (
            self.window.left,
            self.window.top,
            self.window.width,
            self.window.height,
        )

    def capture_page(self, page_num):
        """現在のページをスクリーンショットとして保存します。"""
        region = self.get_window_region()
        screenshot_path = os.path.join(self.temp_dir, f"page_{page_num:04d}.png")

        # スクリーンショット撮影
        screenshot = pyautogui.screenshot(region=region)
        screenshot.save(screenshot_path)
        print(f"Captured page {page_num}")
        return screenshot_path

    def next_page(self):
        """次のページへ移動します。"""
        pyautogui.press("right")
        time.sleep(1.5)  # ページめくりアニメーションを待つ（少し長めに）

    def is_same_image(self, img1_path, img2_path):
        """2つの画像が同じかどうかを判定します。"""
        if not img1_path or not img2_path:
            return False

        try:
            img1 = Image.open(img1_path)
            img2 = Image.open(img2_path)

            if img1.size != img2.size:
                return False

            # 画像の差分を取得
            diff = ImageChops.difference(img1, img2)
            if not diff.getbbox():
                return True  # 完全一致

            return False
        except Exception as e:
            print(f"Error comparing images: {e}")
            return False

    def convert_to_pdf(self, output_filename="output.pdf"):
        """保存された画像をPDFに変換します。"""
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

        with open(output_filename, "wb") as f:
            f.write(img2pdf.convert(image_files))

        print(f"PDF saved to {output_filename}")

    def run(self):
        """変換プロセスを実行します（自動ページ検知）。"""
        try:
            self.locate_window()

            print("Starting in 3 seconds... Please ensure Kindle window is visible.")
            time.sleep(3)

            page_num = 1
            last_screenshot_path = None

            while True:
                current_screenshot_path = self.capture_page(page_num)

                # 前のページと同じかチェック（終了判定）
                if last_screenshot_path and self.is_same_image(
                    last_screenshot_path, current_screenshot_path
                ):
                    print("End of book detected (page content identical to previous).")
                    # 重複した最後の画像を削除
                    os.remove(current_screenshot_path)
                    break

                last_screenshot_path = current_screenshot_path
                self.next_page()
                page_num += 1

                # 安全装置（無限ループ防止）
                if page_num > 2000:
                    print("Safety limit reached (2000 pages). Stopping.")
                    break

            self.convert_to_pdf()
            print("Done!")

        except Exception as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    # テスト用
    # converter = KindleConverter(window_title="メモ帳")
    # converter.run()
    pass
