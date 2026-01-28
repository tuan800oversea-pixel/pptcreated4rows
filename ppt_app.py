import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image

st.title("🖼️ 图片转 PPT 自动排版工具")
st.write("上传多张图片，自动生成 16:9 的四行排版 PPT。")

# 参数设置
SLIDE_WIDTH = Inches(13.333) # 16:9
SLIDE_HEIGHT = Inches(7.5)
TITLE_HEIGHT = Inches(1.2)   # 顶部留白给标题
MARGIN = Inches(0.2)         # 边缘留白
SPACING = Inches(0.1)        # 图片间距
ROW_COUNT = 4                # 固定四行

uploaded_files = st.file_uploader("选择图片文件", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files:
    # 按照文件名排序，确保顺序
    files = sorted(uploaded_files, key=lambda x: x.name)
    
    if st.button("🪄 生成 PPT"):
        prs = Presentation()
        # 设置 16:9 尺寸
        prs.slide_width = SLIDE_WIDTH
        prs.slide_height = SLIDE_HEIGHT
        
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # 使用空白版式
        
        # 计算每行可用高度
        available_height = SLIDE_HEIGHT - TITLE_HEIGHT - (2 * MARGIN) - ((ROW_COUNT - 1) * SPACING)
        row_height = available_height / ROW_COUNT
        
        current_y = TITLE_HEIGHT + MARGIN
        current_x = MARGIN
        
        # 简单的逻辑：平均分配图片到四行
        images_per_row = len(files) // ROW_COUNT + (1 if len(files) % ROW_COUNT > 0 else 0)
        
        for i, file in enumerate(files):
            # 获取图片原始比例
            img_data = Image.open(file)
            orig_w, orig_h = img_data.size
            aspect_ratio = orig_w / orig_h
            
            # 计算在此高度下的等比宽度
            display_width = row_height * aspect_ratio
            
            # 检查是否需要换行（如果超过了 ROW_COUNT 分配的量，或者手动控制）
            if i > 0 and i % images_per_row == 0:
                current_y += row_height + SPACING
                current_x = MARGIN
            
            # 插入图片
            slide.shapes.add_picture(file, current_x, current_y, height=row_height)
            
            current_x += display_width + SPACING

        # 保存并下载
        ppt_buffer = BytesIO()
        prs.save(ppt_buffer)
        
        st.success("🎉 排版完成！")
        st.download_button(
            label="📥 下载 PPT",
            data=ppt_buffer.getvalue(),
            file_name="auto_layout.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
