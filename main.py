import os
import tempfile
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import chardet

def create_photo_presentation(photo_mapping_file, photo_folder_path):
    """Creates a new presentation with slides for each photo and title."""
    prs = Presentation()
    
    # Detect encoding of the uploaded text file
    raw_data = photo_mapping_file.read()
    encoding = chardet.detect(raw_data)['encoding']
    photo_mapping_file.seek(0) # Reset file pointer to beginning after reading for detection
    
    # Read the photo mapping using the detected encoding
    photo_mapping_content = photo_mapping_file.read().decode(encoding)
    lines = photo_mapping_content.strip().split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            continue # Skip empty lines
        if ':' not in line:
            print(f"Warning: Line '{line}' does not contain a colon, skipping.")
            continue # Or handle malformed lines differently

        filename, title = line.split(':', 1) # Split on first colon only
        filename = filename.strip()
        title = title.strip()

        image_path = os.path.join(photo_folder_path, filename)
        if not os.path.exists(image_path):
            print(f"Warning: Image {image_path} not found, skipping.")
            continue # Or handle the missing image as needed

        slide_layout = prs.slide_layouts[5]  # Use a blank layout (typically layout 5)
        slide = prs.slides.add_slide(slide_layout)

        # Add title to the slide
        title_shape = slide.shapes.title
        if title_shape is not None: # Check if title placeholder exists
            title_shape.text = title
            #from pptx.dml.color import RGBColor
            #from pptx.enum.text import PP_ALIGN
            title_shape.text_frame.paragraphs[0].font.size = Pt(24) # Требует from pptx.util import Pt

        # Add image, attempting to fit it well
        # Define margins and maximum space for the image
        left = Inches(0.5)
        top = Inches(1.5) # Leave space for the title
        width = Inches(9)  # Set desired max width
        height = Inches(5) # Set desired max height

        # Calculate aspect ratio to fit image properly without distortion
        from PIL import Image as PILImage
        try:
            with PILImage.open(image_path) as img:
                img_width, img_height = img.size
            
            img_aspect = img_width / img_height
            shape_aspect = width / height

            if img_aspect > shape_aspect:
                # Image is wider relative to its height than the shape -> fit to width
                pic_width = width
                pic_height = int(width / img_aspect)
                pic_left = left
                pic_top = top + (height - pic_height) // 2
            else:
                # Image is taller relative to its width than the shape -> fit to height
                pic_height = height
                pic_width = int(height * img_aspect)
                pic_left = left + (width - pic_width) // 2
                pic_top = top

            slide.shapes.add_picture(image_path, pic_left, pic_top, pic_width, pic_height)
        except Exception as e:
            print(f"Error processing image {image_path}: {e}")
            # Optionally, add a text box indicating the image could not be loaded
            textbox = slide.shapes.add_textbox(left, top, width, height)
            textbox.text = f"Изображение не найдено или невозможно загрузить: {filename}"

    return prs

st.title("Создание отчёта из набора фотографий")

uploaded_mapping_file = st.file_uploader("Загрузите файл подписей к фотографиям (.txt)", type=["txt"])
photo_folder_path = st.text_input("Укажите путь к папке с фотографиями (скопируйте из адресной строки файлового менеджера)")

if st.button("Создать отчёт"):
    if not (uploaded_mapping_file and photo_folder_path):
        st.error("Загрузите файл c подписями к фотографиям и укажите путь к папке с фотографиями.")
    else:
        if not os.path.exists(photo_folder_path):
            st.error(f"Указанный путь к папке с фотографиями не существует: {photo_folder_path}")
        else:
            try:
                # Process the file and create the new presentation
                final_prs = create_photo_presentation(
                    uploaded_mapping_file,
                    photo_folder_path
                )

                # Save the final presentation to a temporary file
                temp_file_path = os.path.join(tempfile.gettempdir(), "photo_report.pptx")
                final_prs.save(temp_file_path)

                # Provide the file for download
                with open(temp_file_path, "rb") as f:
                    st.download_button(
                        label="Скачать отчёт",
                        data=f,
                        file_name="project_photo_report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                # Optional: Clean up the temporary file after download is initiated
                # Note: This might not work perfectly with Streamlit's download button flow
                # os.remove(temp_file_path)

            except Exception as e:
                st.error(f"Произошла ошибка при создании презентации: {e}")
st.markdown(
        """
        <hr>
        <p style="text-align: left; color: gray;">
        <small>
        2025, С.В. Медведев
        </small>
        </p>
        """,
        unsafe_allow_html=True
    )