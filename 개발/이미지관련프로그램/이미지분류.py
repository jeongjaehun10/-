from PIL import Image
import os
import shutil
import random
import tkinter as tk
from tkinter import filedialog
import imagehash

def classify_and_move_images(num_groups=15, num_images_per_folder=15):
    # 이미지 분류 함수
    def classify_images(folder_path, num_groups, progress_var):
        image_files = [f for f in os.listdir(folder_path) if f.endswith(('.png', '.jpg', '.jpeg'))]
        image_hash_map = {}
        
        total_images = len(image_files)
        current_image = 0

        for image_file in image_files:
            image_path = os.path.join(folder_path, image_file)
            image_hash = str(imagehash.average_hash(Image.open(image_path)))
            group = int(image_hash, 16) % num_groups
            
            if group not in image_hash_map:
                image_hash_map[group] = [image_path]
            else:
                image_hash_map[group].append(image_path)

            # Update progress
            current_image += 1
            progress_value = (current_image / total_images) * 100
            progress_var.set(progress_value)
            root.update_idletasks()

        return image_hash_map

    # 이미지 이동 함수
    def move_images(existing_folders, base_folder, output_folder_base, num_images_per_folder, progress_var):
        move_count = 0
        total_folders = len(existing_folders)

        while True:
            current_folder_number = move_count + 1
            new_folder_path = os.path.join(output_folder_base, f'newfolder{current_folder_number}')
            os.makedirs(new_folder_path, exist_ok=True)
            images_moved_count = 0
            min_images_count = min(len([f for f in os.listdir(os.path.join(base_folder, folder))]) for folder in existing_folders)

            for folder_index, folder in enumerate(existing_folders):
                folder_path = os.path.join(base_folder, folder)
                images = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))]

                if images:
                    selected_image = random.choice(images)
                    old_image_path = os.path.join(folder_path, selected_image)
                    new_image_path = os.path.join(new_folder_path, selected_image)
                    shutil.move(old_image_path, new_image_path)
                    images_moved_count += 1

                    if images_moved_count >= num_images_per_folder:
                        break

            if images_moved_count == 0:
                break

            move_count += 1

            if images_moved_count < num_images_per_folder:
                remaining_images_count = num_images_per_folder - images_moved_count
                images_count_dict = {folder: len([f for f in os.listdir(os.path.join(base_folder, folder))]) for folder in existing_folders}
                sorted_folders = sorted(existing_folders, key=lambda folder: images_count_dict[folder], reverse=True)

                for i in range(remaining_images_count):
                    if i < len(sorted_folders):
                        selected_folder = sorted_folders[i]
                        folder_path = os.path.join(base_folder, selected_folder)
                        images = [f for f in os.listdir(folder_path) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))]

                        if images:
                            selected_image = random.choice(images)
                            old_image_path = os.path.join(folder_path, selected_image)
                            new_image_path = os.path.join(new_folder_path, selected_image)
                            shutil.move(old_image_path, new_image_path)

                # Update progress for each folder moved
                progress_value = (move_count / total_folders) * 100
                progress_var.set(progress_value)
                root.update_idletasks()

    # Tkinter를 사용하여 소스 폴더와 분류된 이미지를 이동할 폴더 선택
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select Source and Output Folder")
    root.destroy()

    # Tkinter 프로그레스바 생성
    progress_var = tk.DoubleVar()
    progress_bar = tk.Progressbar(root, variable=progress_var, maximum=100, mode='determinate')
    progress_bar.pack(fill='x', padx=10, pady=10)
    
    # 이미지 분류 및 이동 수행
    image_hash_map = classify_images(folder_path, num_groups, progress_var)
    for group, image_list in image_hash_map.items():
        similar_images_folder = os.path.join(folder_path, f'Similar_Images_Group_{group}')
        os.makedirs(similar_images_folder, exist_ok=True)

        for i, image_path in enumerate(image_list):
            image_name = os.path.basename(image_path)
            new_path = os.path.join(similar_images_folder, f'{image_name}_{i + 1}.jpg')
            shutil.move(image_path, new_path)

    # 이미지 이동
    existing_folders = [folder for folder in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, folder))]
    move_images(existing_folders, folder_path, folder_path, num_images_per_folder, progress_var)

    # 작업 완료 메시지 표시
    tk.messagebox.showinfo("작업 완료", "이미지 분류 및 이동이 완료되었습니다.")
    
    # 프로그램 종료
    root.destroy()

# 함수 호출
classify_and_move_images(num_groups=15, num_images_per_folder=15)
