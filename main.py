from converter import KindleConverter, list_windows


def main():
    print("Kindle to PDF Converter")
    print("-----------------------")

    # ウィンドウ選択
    print("Scanning windows...")
    windows = list_windows()

    if not windows:
        print("No visible windows found.")
        return

    print("\nAvailable windows:")
    for i, w in enumerate(windows):
        print(f"{i}: {w.title}")

    target_window = None
    while True:
        try:
            user_input = input("\nSelect window number (or 'q' to quit): ")
            if user_input.lower() == "q":
                return

            choice = int(user_input)
            if 0 <= choice < len(windows):
                target_window = windows[choice].title
                break
            print("Invalid number. Please try again.")
        except ValueError:
            print("Invalid input. Please enter a number.")

    print(f"\nSelected: {target_window}")

    # オプション選択
    use_ocr = input("Enable OCR (Japanese)? (y/n): ").lower() == "y"

    crop_input = input("Cropping mode? (a: auto, m: manual, n: none) [n]: ").lower()
    crop_mode = "none"
    if crop_input == "a":
        crop_mode = "auto"
    elif crop_input == "m":
        crop_mode = "manual"

    converter = KindleConverter(window_title=target_window)
    converter.run(use_ocr=use_ocr, crop_mode=crop_mode)


if __name__ == "__main__":
    main()
