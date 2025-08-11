import os
import PyPDF2
from PyPDF2 import PageObject, Transformation
from PIL import Image
import io
import tempfile


def remove_pdf_margins(input_path, output_path, zoom_factor=1.2):
    """
    消除PDF白边并放大页面内容

    参数:
    input_path: 输入PDF文件路径
    output_path: 输出PDF文件路径
    zoom_factor: 页面放大倍数 (默认1.2 = 120%)
    """
    try:
        # 检查文件是否存在
        if not os.path.isfile(input_path):
            print(f"错误：文件不存在 - {input_path}")
            return False

        # 创建PDF读写对象
        reader = PyPDF2.PdfReader(input_path)
        writer = PyPDF2.PdfWriter()

        print(f"开始处理: {os.path.basename(input_path)}")
        print(f"总页数: {len(reader.pages)}")

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            print(f"处理中: 第 {page_num + 1}/{len(reader.pages)} 页...")

            # 尝试获取页面尺寸
            try:
                width = float(page.mediabox.width)
                height = float(page.mediabox.height)
            except:
                # 默认A4尺寸 (210x297 mm)
                width, height = 595, 842

            # 创建临时文件处理页面图像
            temp_img_path = None
            try:
                # 尝试将PDF页面转换为图像
                if '/Resources' in page and '/XObject' in page['/Resources']:
                    xObject = page['/Resources']['/XObject'].get_object()
                    for obj in xObject:
                        if xObject[obj]['/Subtype'] == '/Image':
                            temp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                            temp_img_path = temp_img.name
                            temp_img.close()

                            img_data = xObject[obj].get_data()
                            img = Image.open(io.BytesIO(img_data))
                            img.save(temp_img_path)
                            break
            except Exception as e:
                print(f"  警告: 页面图像转换失败 - {str(e)}")

            # 计算裁剪区域
            if temp_img_path and os.path.exists(temp_img_path):
                try:
                    with Image.open(temp_img_path) as img:
                        # 转换为RGB并获取边界框
                        bbox = img.convert("RGB").getbbox()
                        if bbox:
                            # 转换为PDF坐标
                            img_width, img_height = img.size
                            scale_x = width / img_width
                            scale_y = height / img_height

                            # 计算实际边界
                            bbox = (
                                bbox[0] * scale_x,
                                height - bbox[3] * scale_y,  # 注意Y坐标反转
                                bbox[2] * scale_x,
                                height - bbox[1] * scale_y  # 注意Y坐标反转
                            )
                        else:
                            # 使用保守裁剪
                            bbox = (width * 0.05, height * 0.05, width * 0.95, height * 0.95)
                except:
                    bbox = (width * 0.05, height * 0.05, width * 0.95, height * 0.95)

                # 清理临时文件
                os.unlink(temp_img_path)
            else:
                # 使用保守裁剪
                bbox = (width * 0.05, height * 0.05, width * 0.95, height * 0.95)

            # 计算内容尺寸
            content_width = bbox[2] - bbox[0]
            content_height = bbox[3] - bbox[1]

            # 计算缩放和平移参数
            scale = zoom_factor
            new_width = content_width * scale
            new_height = content_height * scale

            # 确保放大后不超过页面尺寸
            if new_width > width:
                scale = width / content_width
            if new_height > height:
                scale = min(scale, height / content_height)

            # 计算居中位置
            tx = (width - content_width * scale) / 2 - bbox[0] * scale
            ty = (height - content_height * scale) / 2 - bbox[1] * scale

            # 创建变换矩阵
            transformation = Transformation().scale(scale, scale).translate(tx, ty)

            # 创建新页面并应用变换
            new_page = PageObject.create_blank_page(width=width, height=height)
            new_page.merge_page(page)
            new_page.add_transformation(transformation)

            # 添加到输出文档
            writer.add_page(new_page)

        # 保存输出文件
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)

        print(f"\n处理完成! 输出文件已保存至: {output_path}")
        print(f"原始文件大小: {os.path.getsize(input_path) / 1024:.1f} KB")
        print(f"新文件大小: {os.path.getsize(output_path) / 1024:.1f} KB")
        return True

    except Exception as e:
        print(f"\n处理过程中出错: {str(e)}")
        return False


def get_valid_file_path(prompt):
    """获取有效的文件路径"""
    while True:
        path = input(prompt).strip()
        if os.path.isfile(path):
            return path
        else:
            print("文件不存在，请重新输入。")


def get_output_path(default_path):
    """获取输出文件路径"""
    path = input(f"输出文件路径 (回车使用默认: {default_path}): ").strip()
    return path if path else default_path


if __name__ == "__main__":
    print("=" * 50)
    print("PDF白边消除工具")
    print("=" * 50)
    print("功能: 自动裁剪PDF白边并将内容放大20%")
    print("提示: 输入PDF文件路径 (绝对路径或相对路径)")

    # 获取输入文件路径
    input_path = get_valid_file_path("\n请输入PDF文件路径: ")

    # 生成默认输出路径
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    default_output = f"{base_name}_no_margins.pdf"

    # 获取输出文件路径
    output_path = get_output_path(default_output)

    # 确认覆盖
    if os.path.exists(output_path):
        overwrite = input(f"文件 {output_path} 已存在，覆盖? (y/n): ").strip().lower()
        if overwrite != 'y':
            print("操作取消。")
            exit()

    # 执行处理
    print("\n开始处理PDF...")
    success = remove_pdf_margins(input_path, output_path)

    if success:
        print("\n操作成功完成!")
    else:
        print("\n处理失败，请检查输入文件。")

    input("\n按回车键退出...")