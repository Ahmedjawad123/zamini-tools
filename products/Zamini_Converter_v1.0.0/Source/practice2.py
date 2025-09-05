import os

def list_first_level(folder_path):
    if not os.path.isdir(folder_path):
        print("Folder does not exist.")
        return

    items = os.listdir(folder_path)
    print("Folders:")
    for item in items:
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):
            print(f"  {item}")

    print("\nFiles:")
    for item in items:
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            print(f"  {item}")

if __name__ == "__main__":
    folder = input("Enter folder path: ")
    list_first_level(folder)
