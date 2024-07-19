import os
from tkinter import filedialog
from tkinter import Tk

def rename_images(folder_path):
    # 폴더 내의 이미지 파일 목록을 가져오기
    image_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]

    for i, image_file in enumerate(image_files):
        # 새로운 파일 이름 생성 (순서대로 1, 2, 3, ...)
        _, extension = os.path.splitext(image_file)
        new_name = f"{i + 1}{extension.lower()}"  # 1부터 시작하는 순서로 수정

        # 이미지 파일 이동 및 이름 변경
        old_path = os.path.join(folder_path, image_file)
        new_path = os.path.join(folder_path, new_name)
        os.rename(old_path, new_path)

def select_folder():
    Tk().withdraw()  # Tkinter 창을 표시하지 않도록 설정
    folder_path = filedialog.askdirectory(title="폴더를 선택하세요.")
    if folder_path:
        rename_images(folder_path)
        print("이미지 파일 이름 변경이 완료되었습니다.")

if __name__ == "__main__":
    select_folder()
