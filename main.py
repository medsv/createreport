import os
import tempfile
import zipfile
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import chardet
from pptx.enum.shapes import PP_PLACEHOLDER_TYPE


def create_photo_presentation(project_title, photo_mapping_file, photo_folder_path):
    """Creates a new presentation with slides for each photo and title."""

    i_title = 0  # Индекс названия проекта
    i_title1 = 12  # Индекс подписи левой фотографии
    i_title2 = 13 # Индекс подписи правой фотографии
    i_page_number = 2 # Индекс номера страницы

    prs = Presentation(os.path.join('template', '04_1.pptx'))
    slide_layout = prs.slide_layouts[0]  # Только один шаблон 
    slide = prs.slides[0]  # Первый слайд
    slide.placeholders[i_title].text = project_title
    # Detect encoding of the uploaded text file
    raw_data = photo_mapping_file.read()
    encoding = chardet.detect(raw_data)['encoding']
    photo_mapping_file.seek(0) # Reset file pointer to beginning after reading for detection
    
    # Read the photo mapping using the detected encoding
    photo_mapping_content = photo_mapping_file.read().decode(encoding)
    lines = photo_mapping_content.strip().split('\n')



    # Координаты областей для вставки фотографий
    left1 = 517206
    left2 = 6235585
    top = 2200942
    width = 5442382
    height = 3996000

    N = 0 # количество вставленных в презинтацию фотографий

    for i,line in enumerate(lines):
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

        # если не нулевая фотогравия и чётная, то создаём слайд
        if i>0 and i % 2 == 0:
            slide = prs.slides.add_slide(slide_layout)
            slide.placeholders[i_title].text = project_title 
            #slide.shapes[i_title].text_frame.text = project_title 
            # slide.shapes[i_page_number].text_frame.text = str(i // 2 + 1) 
                  

        # Add title to the slide
#        title_shape = slide.shapes.title
#        if title_shape is not None: # Check if title placeholder exists
#            title_shape.text = title
#            #from pptx.dml.color import RGBColor
#            #from pptx.enum.text import PP_ALIGN
#            title_shape.text_frame.paragraphs[0].font.size = Pt(24) # Требует from pptx.util import Pt

        # Add image, attempting to fit it well
        # Define margins and maximum space for the image
        #left = Inches(0.5)
        #top = Inches(1.5) # Leave space for the title
        #width = Inches(9)  # Set desired max width
        #height = Inches(5) # Set desired max height

        if i % 2 == 0:
            left = left1
            slide.placeholders[i_title1].text = title
        else:
            left = left2
            slide.placeholders[i_title2].text = title

        N += 1

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
    if N % 2 != 0: # если количество фотографий нечтное, то "обнуляем" заголоаок правой части
        slide.placeholders[i_title2].text = ' '
    return prs

st.title("Создание отчёта из набора фотографий")

project_title: str = st.text_input(label="Введите название проекта", value="Мой проект")
uploaded_zip = st.file_uploader("Загрузите ZIP-архив с фотографиями", type=["zip"])
uploaded_mapping_file = st.file_uploader("Загрузите файл подписей к фотографиям (.txt)", type=["txt"])


if st.button("Создать отчёт"):
    if not (uploaded_mapping_file and uploaded_zip):
        st.error("Загрузите файл с подписями к фотографиям и ZIP-архив с фотографиями.")
    else:
        # Create a temporary directory to extract the ZIP contents
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # Extract the uploaded ZIP file to the temporary directory
                with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Call the presentation creation function with the temporary directory path
                final_prs = create_photo_presentation(
                    project_title,
                    uploaded_mapping_file,
                    temp_dir
                )

                # Save the final presentation to a temporary file
                temp_file_path = os.path.join(tempfile.gettempdir(), "photo_report.pptx")
                final_prs.save(temp_file_path)

                # Provide the file for download
                with open(temp_file_path, "rb") as f:
                    st.download_button(
                        label="Скачать отчёт",
                        data=f,
                        file_name="Фотоотчёт.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    
            except zipfile.BadZipFile:
                st.error("Загруженный файл не является действительным ZIP-архивом.")
            except Exception as e:
                st.error(f"Произошла ошибка при обработке ZIP-архива или создании презентации: {e}")

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