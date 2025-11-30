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
    converter = KindleConverter(window_title=target_window)
    converter.run()


if __name__ == "__main__":
    main()
