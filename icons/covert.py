import aspose.words as aw
import os

def convert_emf_to_png(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith(".emf"):
            emf_path = os.path.join(folder_path, filename)
            png_path = os.path.join(folder_path, filename.replace(".emf", ".png"))
            
            # 使用 Aspose.Words 转换格式
            doc = aw.Document()
            builder = aw.DocumentBuilder(doc)
            shape = builder.insert_image(emf_path)
            
            # 保存为图片
            shape.get_shape_renderer().save(png_path, aw.saving.ImageSaveOptions(aw.SaveFormat.PNG))
            print(f"成功转换: {filename} -> {os.path.basename(png_path)}")

if __name__ == "__main__":
    # 假设当前目录下就是 icons 文件夹
    convert_emf_to_png("./")