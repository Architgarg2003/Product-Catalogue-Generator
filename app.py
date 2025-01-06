# from PIL import Image, ImageDraw
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4
# import pandas as pd
# import textwrap

# # Load the Excel file
# file_path = "Tools List.xlsx"
# data = pd.read_excel(file_path)

# # Placeholder image for products
# placeholder_image_path = "placeholder.jpg"

# # Create a simple placeholder image
# placeholder = Image.new("RGB", (300, 300), color="lightgray")
# draw = ImageDraw.Draw(placeholder)
# draw.text((100, 140), "Image", fill="black")
# placeholder.save(placeholder_image_path)

# # Page settings
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_COLS = 3
# GRID_ROWS = 4
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = (PAGE_HEIGHT - 100) / GRID_ROWS  # Leave space for the header

# # Font settings
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 20
# font_size_address = 12
# font_size_name = 8
# font_size_price = 10

# # Output PDF file
# output_pdf_path = "updated_layout_output.pdf"


# def wrap_text(text, width):
#     """Wrap text to fit within a specific width."""
#     return textwrap.fill(text, width)


# def create_pdf(data, placeholder_image, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)

#     # Add heading and address to the first page
#     c.setFont(font_heading, font_size_heading)
#     c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")
#     c.setFont(font_text, font_size_address)
#     c.drawCentredString(
#         PAGE_WIDTH / 2, PAGE_HEIGHT - 80, "Address: 123 Example Street, City Name, State, Country"
#     )

#     current_row = 0
#     current_col = 0

#     for _, row in data.iterrows():
#         # Check if starting a new page
#         if current_row == 0 and current_col == 0 and c._pageNumber > 1:
#             c.showPage()

#         # Get product details
#         product_name = str(row.get("Name", "Unknown Product"))
#         price = str(row.get("Price E", "Unknown Price"))

#         # Calculate cell position
#         x = MARGIN + current_col * CELL_WIDTH
#         y = PAGE_HEIGHT - 100 - (current_row + 1) * CELL_HEIGHT

#         # Draw image
#         try:
#             img = Image.open(placeholder_image)
#             img.thumbnail((CELL_WIDTH - 20, CELL_HEIGHT // 2))
#             img.save("temp_image.jpg")
#             c.drawImage(
#                 "temp_image.jpg",
#                 x + (CELL_WIDTH - img.width) / 2,
#                 y + CELL_HEIGHT - img.height - 10,
#                 width=img.width,
#                 height=img.height,
#             )
#         except Exception as e:
#             print(f"Error loading image: {e}")

#         # Draw product name and wrap text
#         c.setFont(font_text, font_size_name)
#         wrapped_name = wrap_text(product_name, 30)  # Adjust wrapping width
#         text_y = y + CELL_HEIGHT - img.height - 25
#         for line in wrapped_name.splitlines():
#             c.drawString(x + 5, text_y, line)
#             text_y -= 10  # Adjust line spacing

#         # Draw price
#         c.setFont(font_text, font_size_price)
#         c.drawString(x + 5, y + 20, f"₹{price}")

#         # Move to the next cell
#         current_col += 1
#         if current_col >= GRID_COLS:
#             current_col = 0
#             current_row += 1
#             if current_row >= GRID_ROWS:
#                 c.showPage()
#                 current_row = 0

#     c.save()


# # Generate the PDF
# create_pdf(data, placeholder_image_path, output_pdf_path)

# print(f"PDF generated successfully: {output_pdf_path}")




# from PIL import Image, ImageDraw
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4
# import pandas as pd
# import textwrap

# # Load the Excel file
# file_path = "Tools List.xlsx"
# data = pd.read_excel(file_path)

# # Placeholder image for products
# placeholder_image_path = "placeholder.jpg"

# # Create a simple placeholder image
# placeholder = Image.new("RGB", (300, 300), color="lightgray")
# draw = ImageDraw.Draw(placeholder)
# draw.text((100, 140), "Image", fill="black")
# placeholder.save(placeholder_image_path)

# # Page settings
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_COLS = 3
# GRID_ROWS = 4
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = (PAGE_HEIGHT - 100) / GRID_ROWS  # Leave space for the header

# # Font settings
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 20
# font_size_address = 12
# font_size_name = 8
# font_size_price = 10

# # Output PDF file
# output_pdf_path = "updated_layout_output.pdf"

# def wrap_text(text, width):
#     """Wrap text to fit within a specific width."""
#     return textwrap.fill(text, width)

# def create_pdf(data, placeholder_image, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)
#     items_per_page = GRID_ROWS * GRID_COLS
#     total_items = len(data)
    
#     for page_num in range((total_items + items_per_page - 1) // items_per_page):
#         if page_num > 0:
#             c.showPage()
            
#         # Add heading and address to each page
#         c.setFont(font_heading, font_size_heading)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")
#         c.setFont(font_text, font_size_address)
#         c.drawCentredString(
#             PAGE_WIDTH / 2, 
#             PAGE_HEIGHT - 80, 
#             "Address: 123 Example Street, City Name, State, Country"
#         )
        
#         # Process items for current page
#         start_idx = page_num * items_per_page
#         end_idx = min((page_num + 1) * items_per_page, total_items)
        
#         for idx in range(start_idx, end_idx):
#             row = data.iloc[idx]
#             current_row = (idx - start_idx) // GRID_COLS
#             current_col = (idx - start_idx) % GRID_COLS
            
#             # Get product details
#             product_name = str(row.get("Name", "Unknown Product"))
#             price = str(row.get("Price E", "Unknown Price"))
            
#             # Calculate cell position
#             x = MARGIN + current_col * CELL_WIDTH
#             y = PAGE_HEIGHT - 100 - (current_row + 1) * CELL_HEIGHT
            
#             # Draw image
#             try:
#                 img = Image.open(placeholder_image)
#                 img.thumbnail((CELL_WIDTH - 20, CELL_HEIGHT // 2))
#                 img.save("temp_image.jpg")
#                 c.drawImage(
#                     "temp_image.jpg",
#                     x + (CELL_WIDTH - img.width) / 2,
#                     y + CELL_HEIGHT - img.height - 10,
#                     width=img.width,
#                     height=img.height,
#                 )
#             except Exception as e:
#                 print(f"Error loading image: {e}")
            
#             # Draw product name and wrap text
#             c.setFont(font_text, font_size_name)
#             wrapped_name = wrap_text(product_name, 30)  # Adjust wrapping width
#             text_y = y + CELL_HEIGHT - img.height - 25
#             for line in wrapped_name.splitlines():
#                 c.drawString(x + 5, text_y, line)
#                 text_y -= 10  # Adjust line spacing
            
#             # Draw price
#             c.setFont(font_text, font_size_price)
#             c.drawString(x + 5, y + 20, f"₹{price}")
    
#     c.save()

# # Generate the PDF
# create_pdf(data, placeholder_image_path, output_pdf_path)
# print(f"PDF generated successfully: {output_pdf_path}")


# '''grid lines'''



# from PIL import Image, ImageDraw
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4
# import pandas as pd
# import textwrap

# # Load the Excel file
# file_path = "Tools List.xlsx"
# data = pd.read_excel(file_path)

# # Placeholder image for products
# placeholder_image_path = "placeholder.jpg"

# # Create a simple placeholder image
# placeholder = Image.new("RGB", (300, 300), color="lightgray")
# draw = ImageDraw.Draw(placeholder)
# draw.text((100, 140), "Image", fill="black")
# placeholder.save(placeholder_image_path)

# # Page settings
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_COLS = 3
# GRID_ROWS = 4
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = (PAGE_HEIGHT - 100) / GRID_ROWS  # Leave space for the header

# # Font settings
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 20
# font_size_address = 12
# font_size_name = 8
# font_size_price = 10

# # Output PDF file
# output_pdf_path = "updated_layout_output.pdf"

# def wrap_text(text, width):
#     """Wrap text to fit within a specific width."""
#     return textwrap.fill(text, width)

# def draw_grid(canvas_obj, start_y):
#     """Draw grid lines for the product layout."""
#     # Draw horizontal lines
#     for row in range(GRID_ROWS + 1):
#         y = start_y - (row * CELL_HEIGHT)
#         canvas_obj.line(MARGIN, y, PAGE_WIDTH - MARGIN, y)

#     # Draw vertical lines
#     for col in range(GRID_COLS + 1):
#         x = MARGIN + (col * CELL_WIDTH)
#         canvas_obj.line(x, start_y, x, start_y - (GRID_ROWS * CELL_HEIGHT))

# def create_pdf(data, placeholder_image, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)
#     items_per_page = GRID_ROWS * GRID_COLS
#     total_items = len(data)
    
#     for page_num in range((total_items + items_per_page - 1) // items_per_page):
#         if page_num > 0:
#             c.showPage()
            
#         # Add heading and address to each page
#         c.setFont(font_heading, font_size_heading)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")
#         c.setFont(font_text, font_size_address)
#         c.drawCentredString(
#             PAGE_WIDTH / 2, 
#             PAGE_HEIGHT - 80, 
#             "Address: 123 Example Street, City Name, State, Country"
#         )
        
#         # Draw grid lines
#         grid_start_y = PAGE_HEIGHT - 100
#         draw_grid(c, grid_start_y)
        
#         # Process items for current page
#         start_idx = page_num * items_per_page
#         end_idx = min((page_num + 1) * items_per_page, total_items)
        
#         for idx in range(start_idx, end_idx):
#             row = data.iloc[idx]
#             current_row = (idx - start_idx) // GRID_COLS
#             current_col = (idx - start_idx) % GRID_COLS
            
#             # Get product details
#             product_name = str(row.get("Name", "Unknown Product"))
#             price = str(row.get("Price E", "Unknown Price"))
            
#             # Calculate cell position
#             x = MARGIN + current_col * CELL_WIDTH
#             y = PAGE_HEIGHT - 100 - (current_row + 1) * CELL_HEIGHT
            
#             # Draw image
#             try:
#                 img = Image.open(placeholder_image)
#                 img.thumbnail((CELL_WIDTH - 20, CELL_HEIGHT // 2))
#                 img.save("temp_image.jpg")
#                 c.drawImage(
#                     "temp_image.jpg",
#                     x + (CELL_WIDTH - img.width) / 2,
#                     y + CELL_HEIGHT - img.height - 10,
#                     width=img.width,
#                     height=img.height,
#                 )
#             except Exception as e:
#                 print(f"Error loading image: {e}")
            
#             # Draw product name and wrap text
#             c.setFont(font_text, font_size_name)
#             wrapped_name = wrap_text(product_name, 30)  # Adjust wrapping width
#             text_y = y + CELL_HEIGHT - img.height - 25
#             for line in wrapped_name.splitlines():
#                 c.drawString(x + 5, text_y, line)
#                 text_y -= 10  # Adjust line spacing
            
#             # Draw price
#             c.setFont(font_text, font_size_price)
#             c.drawString(x + 5, y + 20, f"₹{price}")
    
#     c.save()

# # Generate the PDF
# create_pdf(data, placeholder_image_path, output_pdf_path)
# print(f"PDF generated successfully: {output_pdf_path}")



# '''final w grid lines'''



# from PIL import Image, ImageDraw
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4
# import pandas as pd
# import textwrap

# # Load the Excel file
# file_path = "Tools List.xlsx"
# data = pd.read_excel(file_path)

# # Placeholder image for products
# placeholder_image_path = "placeholder.jpg"

# # Create a simple placeholder image
# placeholder = Image.new("RGB", (300, 300), color="lightgray")
# draw = ImageDraw.Draw(placeholder)
# draw.text((100, 140), "Image", fill="black")
# placeholder.save(placeholder_image_path)

# # Page settings
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_COLS = 3
# GRID_ROWS = 4
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = (PAGE_HEIGHT - 100) / GRID_ROWS  # Leave space for the header

# # Image settings
# IMAGE_HEIGHT = CELL_HEIGHT * 0.7  # Increased to 70% of cell height
# IMAGE_WIDTH = CELL_WIDTH - 20

# # Font settings
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 20
# font_size_address = 12
# font_size_name = 8
# font_size_price = 10

# # Output PDF file
# output_pdf_path = "updated_layout_output.pdf"

# def wrap_text(text, width):
#     """Wrap text to fit within a specific width."""
#     return textwrap.fill(text, width)

# def draw_grid(canvas_obj, start_y):
#     """Draw grid lines for the product layout."""
#     # Draw horizontal lines
#     for row in range(GRID_ROWS + 1):
#         y = start_y - (row * CELL_HEIGHT)
#         canvas_obj.line(MARGIN, y, PAGE_WIDTH - MARGIN, y)

#     # Draw vertical lines
#     for col in range(GRID_COLS + 1):
#         x = MARGIN + (col * CELL_WIDTH)
#         canvas_obj.line(x, start_y, x, start_y - (GRID_ROWS * CELL_HEIGHT))

# def create_pdf(data, placeholder_image, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)
#     items_per_page = GRID_ROWS * GRID_COLS
#     total_items = len(data)
    
#     for page_num in range((total_items + items_per_page - 1) // items_per_page):
#         if page_num > 0:
#             c.showPage()
            
#         # Add heading and address to each page
#         c.setFont(font_heading, font_size_heading)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")
#         c.setFont(font_text, font_size_address)
#         c.drawCentredString(
#             PAGE_WIDTH / 2, 
#             PAGE_HEIGHT - 80, 
#             "Address: 123 Example Street, City Name, State, Country"
#         )
        
#         # Draw grid lines
#         grid_start_y = PAGE_HEIGHT - 100
#         draw_grid(c, grid_start_y)
        
#         # Process items for current page
#         start_idx = page_num * items_per_page
#         end_idx = min((page_num + 1) * items_per_page, total_items)
        
#         for idx in range(start_idx, end_idx):
#             row = data.iloc[idx]
#             current_row = (idx - start_idx) // GRID_COLS
#             current_col = (idx - start_idx) % GRID_COLS
            
#             # Get product details
#             product_name = str(row.get("Name", "Unknown Product"))
#             price = str(row.get("Price E", "Unknown Price"))
            
#             # Calculate cell position
#             x = MARGIN + current_col * CELL_WIDTH
#             y = PAGE_HEIGHT - 100 - (current_row + 1) * CELL_HEIGHT
            
#             # Draw image
#             try:
#                 img = Image.open(placeholder_image)
#                 # Set thumbnail size to match our desired dimensions
#                 img.thumbnail((IMAGE_WIDTH, IMAGE_HEIGHT))
#                 img.save("temp_image.jpg")
#                 c.drawImage(
#                     "temp_image.jpg",
#                     x + (CELL_WIDTH - img.width) / 2,
#                     y + CELL_HEIGHT - img.height - 5,  # Reduced top padding
#                     width=img.width,
#                     height=img.height,
#                 )
#             except Exception as e:
#                 print(f"Error loading image: {e}")
            
#             # Draw product name and wrap text
#             c.setFont(font_text, font_size_name)
#             wrapped_name = wrap_text(product_name, 30)  # Adjust wrapping width
#             text_y = y + CELL_HEIGHT - IMAGE_HEIGHT - 15  # Adjusted position
#             for line in wrapped_name.splitlines():
#                 c.drawString(x + 5, text_y, line)
#                 text_y -= 10  # Adjust line spacing
            
#             # Draw price with rupee symbol
#             c.setFont(font_text, font_size_price)
#             c.drawString(x + 5, text_y - 5, f"₹ {price}")  # Reduced gap and added rupee symbol
    
#     c.save()

# # Generate the PDF
# create_pdf(data, placeholder_image_path, output_pdf_path)
# print(f"PDF generated successfully: {output_pdf_path}")



# '''style Final'''


# from reportlab.lib import colors
# from reportlab.lib.pagesizes import A4
# from reportlab.pdfgen import canvas
# from PIL import Image, ImageDraw
# import pandas as pd
# import textwrap

# # Page and grid settings
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_COLS = 3
# GRID_ROWS = 4
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = (PAGE_HEIGHT - 120) / GRID_ROWS  # Adjusted for header
# IMAGE_HEIGHT = CELL_HEIGHT * 0.6
# IMAGE_WIDTH = CELL_WIDTH - 20

# # Fonts and colors
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 24
# font_size_address = 10
# font_size_name = 10
# font_size_price = 12
# color_heading = colors.HexColor("#2C3E50")
# color_text = colors.HexColor("#34495E")
# color_price = colors.HexColor("#27AE60")
# bg_color_cell = colors.HexColor("#ECF0F1")
# border_color_cell = colors.HexColor("#BDC3C7")

# # Placeholder image
# placeholder_image_path = "placeholder.jpg"
# placeholder = Image.new("RGB", (300, 300), color="lightgray")
# draw = ImageDraw.Draw(placeholder)
# draw.text((100, 140), "Image", fill="black")
# placeholder.save(placeholder_image_path)

# # Function to wrap text
# def wrap_text(text, width):
#     return textwrap.fill(text, width)

# # Function to draw the grid with enhanced styling
# def draw_grid(canvas_obj, start_y):
#     for row in range(GRID_ROWS):
#         for col in range(GRID_COLS):
#             x = MARGIN + col * CELL_WIDTH
#             y = start_y - row * CELL_HEIGHT
#             # Draw cell background
#             canvas_obj.setFillColor(bg_color_cell)
#             canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=True, stroke=False)
#             # Draw cell border
#             canvas_obj.setStrokeColor(border_color_cell)
#             canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=False, stroke=True)

# # Function to create the PDF
# def create_pdf(data, placeholder_image, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)
#     items_per_page = GRID_ROWS * GRID_COLS
#     total_items = len(data)

#     for page_num in range((total_items + items_per_page - 1) // items_per_page):
#         if page_num > 0:
#             c.showPage()

#         # Add heading
#         c.setFont(font_heading, font_size_heading)
#         c.setFillColor(color_heading)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")

#         # Add address
#         c.setFont(font_text, font_size_address)
#         c.setFillColor(color_text)
#         c.drawCentredString(
#             PAGE_WIDTH / 2,
#             PAGE_HEIGHT - 70,
#             "Address: 123 Example Street, City Name, State, Country"
#         )

#         # Draw grid
#         grid_start_y = PAGE_HEIGHT - 100
#         draw_grid(c, grid_start_y)

#         # Process items for the current page
#         start_idx = page_num * items_per_page
#         end_idx = min((page_num + 1) * items_per_page, total_items)

#         for idx in range(start_idx, end_idx):
#             row = data.iloc[idx]
#             current_row = (idx - start_idx) // GRID_COLS
#             current_col = (idx - start_idx) % GRID_COLS

#             # Get product details
#             product_name = str(row.get("Name", "Unknown Product"))
#             price = str(row.get("Price E", "Unknown Price"))

#             # Calculate cell position
#             x = MARGIN + current_col * CELL_WIDTH
#             y = grid_start_y - current_row * CELL_HEIGHT

#             # Draw image
#             try:
#                 img = Image.open(placeholder_image)
#                 img.thumbnail((IMAGE_WIDTH, IMAGE_HEIGHT))
#                 img.save("temp_image.jpg")
#                 c.drawImage(
#                     "temp_image.jpg",
#                     x + (CELL_WIDTH - IMAGE_WIDTH) / 2,
#                     y - IMAGE_HEIGHT - 10,
#                     width=IMAGE_WIDTH,
#                     height=IMAGE_HEIGHT,
#                 )
#             except Exception as e:
#                 print(f"Error loading image: {e}")

#             # Draw product name
#             c.setFont(font_text, font_size_name)
#             c.setFillColor(color_text)
#             wrapped_name = wrap_text(product_name, 25)
#             text_y = y - IMAGE_HEIGHT - 20
#             for line in wrapped_name.splitlines():
#                 c.drawString(x + 10, text_y, line)
#                 text_y -= 12

#             # Draw price
#             c.setFont(font_text, font_size_price)
#             c.setFillColor(color_price)
#             c.drawString(x + 10, text_y - 5, f"₹ {price}")

#     c.save()
#     print(f"PDF generated successfully: {output_pdf}")

# # Generate the PDF
# file_path = "Tools List.xlsx"
# data = pd.read_excel(file_path)
# output_pdf_path = "stylish_layout_output.pdf"
# create_pdf(data, placeholder_image_path, output_pdf_path)


# '''image'''






# import os
# import pandas as pd
# from openpyxl import load_workbook
# from openpyxl.drawing.image import Image as XLImage
# from PIL import Image
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import A4
# from reportlab.lib.colors import black, gray, blue

# # Constants
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_ROWS = 4
# GRID_COLS = 3
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = 200
# IMAGE_WIDTH = 100
# IMAGE_HEIGHT = 100

# # Font and colors
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 18
# font_size_name = 10
# font_size_price = 10
# font_size_address = 8
# color_heading = blue
# color_text = black
# color_price = gray

# import zipfile
# from PIL import Image

# def extract_images_with_zip(file_path, data):
#     """
#     Extract embedded images from an Excel file using the ZIP archive structure.
#     """
#     images = []
#     image_dir = 'temp_images'
    
#     # Create a directory for temp images if it doesn't exist
#     if not os.path.exists(image_dir):
#         os.makedirs(image_dir)

#     try:
#         with zipfile.ZipFile(file_path, 'r') as z:
#             # Look for images in the "xl/media" folder
#             image_files = [f for f in z.namelist() if f.startswith('xl/media/')]
#             for i, image_file in enumerate(image_files):
#                 temp_path = os.path.join(image_dir, f"image_{i + 1}.png")
#                 with z.open(image_file) as img_data:
#                     with open(temp_path, 'wb') as f:
#                         f.write(img_data.read())
#                 images.append(temp_path)

#         # Ensure we have the same number of images as rows in the DataFrame
#         if len(images) < len(data):
#             placeholder_path = os.path.join(image_dir, 'placeholder.png')
#             if not os.path.exists(placeholder_path):
#                 placeholder = Image.new('RGB', (100, 100), color='lightgray')
#                 placeholder.save(placeholder_path)
#             images.extend([placeholder_path] * (len(data) - len(images)))
#     except Exception as e:
#         print(f"Error extracting images with zip: {e}")
#         images = [os.path.join(image_dir, 'placeholder.png')] * len(data)

#     print(f"Extracted {len(images)} images from the Excel file.")
#     return images


# def clean_up_temp_images():
#     if os.path.exists('temp_images'):
#         for file in os.listdir('temp_images'):
#             try:
#                 os.remove(os.path.join('temp_images', file))
#             except Exception as e:
#                 print(f"Error removing temp file {file}: {e}")
#         os.rmdir('temp_images')

# def draw_grid(canvas, start_y):
#     canvas.setStrokeColor(black)
#     canvas.setLineWidth(0.5)
#     for row in range(GRID_ROWS + 1):
#         y = start_y - row * CELL_HEIGHT
#         canvas.line(MARGIN, y, PAGE_WIDTH - MARGIN, y)
#     for col in range(GRID_COLS + 1):
#         x = MARGIN + col * CELL_WIDTH
#         canvas.line(x, start_y, x, start_y - GRID_ROWS * CELL_HEIGHT)

# def wrap_text(text, max_chars):
#     words = str(text).split()
#     lines = []
#     line = []
#     for word in words:
#         if len(" ".join(line + [word])) <= max_chars:
#             line.append(word)
#         else:
#             lines.append(" ".join(line))
#             line = [word]
#     if line:
#         lines.append(" ".join(line))
#     return "\n".join(lines)

# def create_pdf_with_images(data, image_paths, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)
#     items_per_page = GRID_ROWS * GRID_COLS
#     total_items = len(data)
#     print(f"Creating PDF with {total_items} items...")

#     for page_num in range((total_items + items_per_page - 1) // items_per_page):
#         if page_num > 0:
#             c.showPage()
        
#         print(f"Processing page {page_num + 1}...")

#         # Add heading and address
#         c.setFont(font_heading, font_size_heading)
#         c.setFillColor(color_heading)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")
#         c.setFont(font_text, font_size_address)
#         c.setFillColor(color_text)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 70, "Address: 123 Example Street, City Name, State, Country")

#         # Draw grid
#         grid_start_y = PAGE_HEIGHT - 100
#         draw_grid(c, grid_start_y)

#         # Process items for the current page
#         start_idx = page_num * items_per_page
#         end_idx = min((page_num + 1) * items_per_page, total_items)

#         for idx in range(start_idx, end_idx):
#             row = data.iloc[idx]
#             current_row = (idx - start_idx) // GRID_COLS
#             current_col = (idx - start_idx) % GRID_COLS

#             product_name = str(row.get("Name", "Unknown Product"))
#             price = str(row.get("Price E", "Unknown Price"))
#             image_path = image_paths[idx] if idx < len(image_paths) else None

#             x = MARGIN + current_col * CELL_WIDTH
#             y = grid_start_y - current_row * CELL_HEIGHT

#             # Draw image
#             if image_path and os.path.exists(image_path):
#                 try:
#                     c.drawImage(image_path, x + (CELL_WIDTH - IMAGE_WIDTH) / 2, 
#                               y - IMAGE_HEIGHT - 10, width=IMAGE_WIDTH, height=IMAGE_HEIGHT)
#                 except Exception as e:
#                     print(f"Error adding image to PDF for {product_name}: {e}")

#             # Draw product name
#             c.setFont(font_text, font_size_name)
#             c.setFillColor(color_text)
#             wrapped_name = wrap_text(product_name, 25)
#             text_y = y - IMAGE_HEIGHT - 20
#             for line in wrapped_name.splitlines():
#                 c.drawString(x + 10, text_y, line)
#                 text_y -= 12

#             # Draw price
#             c.setFont(font_text, font_size_price)
#             c.setFillColor(color_price)
#             c.drawString(x + 10, text_y - 5, f"₹ {price}")

#         print(f"Completed page {page_num + 1}")

#     c.save()
#     print(f"PDF generated successfully: {output_pdf}")

# def main():
#     try:
#         excel_file_path = "Tools List.xlsx"
#         print("Reading Excel file...")
#         data = pd.read_excel(excel_file_path)
#         print(f"Found {len(data)} items in the Excel file")
        
#         print("Extracting images...")
#         image_paths = extract_images_with_zip(excel_file_path, data)
#         print(f"Found {len(image_paths)} images")
        
#         output_pdf_path = "stylish_layout_with_images.pdf"
#         print("Generating PDF...")
#         create_pdf_with_images(data, image_paths, output_pdf_path)
#         print("Process completed successfully!")
#     except Exception as e:
#         print(f"An error occurred: {e}")
#     finally:
#         clean_up_temp_images()

# if __name__ == "__main__":
#     main()



# '''image order'''


# import os
# import pandas as pd
# from openpyxl import load_workbook
# from openpyxl.drawing.image import Image as XLImage
# import io

# def extract_images_and_add_column(excel_file, output_excel_file):
#     # Step 1: Extract images from Excel
#     image_dir = "extracted_images"
#     if not os.path.exists(image_dir):
#         os.makedirs(image_dir)

#     images = []
#     workbook = load_workbook(excel_file)
#     sheet = workbook.active

#     # Extract images and store them with their respective row indices
#     for idx, image in enumerate(sheet._images):
#         if isinstance(image, XLImage):
#             image_name = f"image_{idx + 1}.png"
#             image_path = os.path.join(image_dir, image_name)
#             # Save the image to disk using the image's reference stream
#             img_stream = io.BytesIO(image._data())  # Get the image data as bytes
#             with open(image_path, "wb") as f:
#                 f.write(img_stream.getvalue())  # Write bytes to file
#             images.append((idx, image_name))  # Store index and name

#     print(f"Extracted {len(images)} images and saved to '{image_dir}'")

#     # Step 2: Map images to rows and add "Image Name" column
#     data = pd.DataFrame(sheet.values)
#     headers = data.iloc[0]  # Assuming the first row is the header
#     data = data[1:]  # Exclude header row
#     data.columns = headers

#     # Add "Image Name" column
#     image_column = ["" for _ in range(len(data))]
    
#     for idx, image_name in images:
#         if idx < len(image_column):
#             image_column[idx] = image_name

#     data["Image Name"] = image_column

#     # Save the updated Excel with the new column
#     data.to_excel(output_excel_file, index=False)
#     print(f"Updated Excel file saved to '{output_excel_file}' with 'Image Name' column.")

# # Usage
# excel_file = "test.xlsx"  # Input Excel file with images
# output_excel_file = "updated_test.xlsx"  # Output Excel file with "Image Name" column
# extract_images_and_add_column(excel_file, output_excel_file)






# import os
# import io
# import shutil
# import pandas as pd
# from PIL import Image, ImageDraw
# from openpyxl import load_workbook
# from openpyxl.drawing.image import Image as XLImage
# from reportlab.lib import colors
# from reportlab.lib.pagesizes import A4
# from reportlab.pdfgen import canvas
# import textwrap

# # Page and grid settings
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_COLS = 3
# GRID_ROWS = 4
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = (PAGE_HEIGHT - 120) / GRID_ROWS  # Adjusted for header
# IMAGE_HEIGHT = CELL_HEIGHT * 0.6
# IMAGE_WIDTH = CELL_WIDTH - 20

# # Fonts and colors
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 24
# font_size_address = 10
# font_size_name = 10
# font_size_price = 12
# color_heading = colors.HexColor("#2C3E50")
# color_text = colors.HexColor("#34495E")
# color_price = colors.HexColor("#27AE60")
# bg_color_cell = colors.HexColor("#ECF0F1")
# border_color_cell = colors.HexColor("#BDC3C7")

# # Function to wrap text
# def wrap_text(text, width):
#     return textwrap.fill(text, width)

# # Function to draw the grid with enhanced styling
# def draw_grid(canvas_obj, start_y):
#     for row in range(GRID_ROWS):
#         for col in range(GRID_COLS):
#             x = MARGIN + col * CELL_WIDTH
#             y = start_y - row * CELL_HEIGHT
#             # Draw cell background
#             canvas_obj.setFillColor(bg_color_cell)
#             canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=True, stroke=False)
#             # Draw cell border
#             canvas_obj.setStrokeColor(border_color_cell)
#             canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=False, stroke=True)

# # Function to create the PDF
# def create_pdf(data, placeholder_image, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)
#     items_per_page = GRID_ROWS * GRID_COLS
#     total_items = len(data)

#     for page_num in range((total_items + items_per_page - 1) // items_per_page):
#         if page_num > 0:
#             c.showPage()

#         # Add heading
#         c.setFont(font_heading, font_size_heading)
#         c.setFillColor(color_heading)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")

#         # Add address
#         c.setFont(font_text, font_size_address)
#         c.setFillColor(color_text)
#         c.drawCentredString(
#             PAGE_WIDTH / 2,
#             PAGE_HEIGHT - 70,
#             "Address: 123 Example Street, City Name, State, Country"
#         )

#         # Draw grid
#         grid_start_y = PAGE_HEIGHT - 100
#         draw_grid(c, grid_start_y)

#         # Process items for the current page
#         start_idx = page_num * items_per_page
#         end_idx = min((page_num + 1) * items_per_page, total_items)

#         for idx in range(start_idx, end_idx):
#             row = data.iloc[idx]
#             current_row = (idx - start_idx) // GRID_COLS
#             current_col = (idx - start_idx) % GRID_COLS

#             # Get product details
#             product_name = str(row.get("Name", "Unknown Product"))
#             price = str(row.get("Price E", "Unknown Price"))
#             image_name = row.get("Image Name", None)

#             # Calculate cell position
#             x = MARGIN + current_col * CELL_WIDTH
#             y = grid_start_y - current_row * CELL_HEIGHT

#             # Draw image
#             # Draw image
#             try:
#                 # Construct the full image path
#                 image_path = os.path.join("extracted_images", image_name) if image_name else placeholder_image

#                 # Check if the image exists
#                 if os.path.exists(image_path):
#                     img = Image.open(image_path)
#                 else:
#                     print(f"Image not found: {image_path}, using placeholder.")
#                     img = Image.open(placeholder_image)

#                 # Convert and resize the image
#                 img = img.convert("RGB")
#                 img.thumbnail((IMAGE_WIDTH, IMAGE_HEIGHT))

#                 # Save to a temporary file
#                 temp_image_path = f"temp_image_{idx}.jpg"
#                 img.save(temp_image_path)

#                 # Draw the image from the temporary file path
#                 c.drawImage(
#                     temp_image_path,
#                     x + (CELL_WIDTH - IMAGE_WIDTH) / 2,
#                     y - IMAGE_HEIGHT - 10,
#                     width=IMAGE_WIDTH,
#                     height=IMAGE_HEIGHT,
#                 )

#                 # Cleanup the temporary file after usage
#                 os.remove(temp_image_path)
#             except Exception as e:
#                 print(f"Error loading image for {product_name}: {e}")

#             # Draw product name
#             c.setFont(font_text, font_size_name)
#             c.setFillColor(color_text)
#             wrapped_name = wrap_text(product_name, 25)
#             text_y = y - IMAGE_HEIGHT - 20
#             for line in wrapped_name.splitlines():
#                 c.drawString(x + 10, text_y, line)
#                 text_y -= 12

#             # Draw price
#             c.setFont(font_text, font_size_price)
#             c.setFillColor(color_price)
#             c.drawString(x + 10, text_y - 5, f"₹ {price}")

#     c.save()
#     print(f"PDF generated successfully: {output_pdf}")

# # Function to extract images and add the "Image Name" column
# def extract_images_and_add_column(excel_file, output_excel_file):
#     image_dir = "extracted_images"
#     if not os.path.exists(image_dir):
#         os.makedirs(image_dir)

#     images = []
#     workbook = load_workbook(excel_file)
#     sheet = workbook.active

#     for idx, image in enumerate(sheet._images):
#         if isinstance(image, XLImage):
#             image_name = f"image_{idx + 1}.png"
#             image_path = os.path.join(image_dir, image_name)
#             img_stream = io.BytesIO(image._data())  # Get the image data as bytes
#             with open(image_path, "wb") as f:
#                 f.write(img_stream.getvalue())  # Write bytes to file
#             images.append((idx, image_name))

#     print(f"Extracted {len(images)} images and saved to '{image_dir}'")

#     data = pd.DataFrame(sheet.values)
#     headers = data.iloc[0]  # Assuming the first row is the header
#     data = data[1:]  # Exclude header row
#     data.columns = headers

#     image_column = ["" for _ in range(len(data))]
#     for idx, image_name in images:
#         if idx < len(image_column):
#             image_column[idx] = image_name

#     data["Image Name"] = image_column
#     data.to_excel(output_excel_file, index=False)
#     print(f"Updated Excel file saved to '{output_excel_file}' with 'Image Name' column.")

# # Main Script
# excel_file = "Tools List-2.xlsx"
# updated_excel_file = "updated_test.xlsx"
# placeholder_image_path = "placeholder.jpg"
# output_pdf_path = "stylish_layout_output.pdf"

# # Create a placeholder image
# placeholder = Image.new("RGB", (300, 300), color="lightgray")
# draw = ImageDraw.Draw(placeholder)
# draw.text((100, 140), "Image", fill="black")
# placeholder.save(placeholder_image_path)

# # Extract images and add the "Image Name" column
# extract_images_and_add_column(excel_file, updated_excel_file)

# # Generate the PDF
# data = pd.read_excel(updated_excel_file)
# create_pdf(data, placeholder_image_path, output_pdf_path)

# # Remove the extracted_images folder
# # shutil.rmtree("extracted_images", ignore_errors=True)


# import os
# from openpyxl import load_workbook
# from openpyxl_image_loader import SheetImageLoader
# from PIL import Image, ImageDraw
# import shutil

# # Function to create a placeholder image
# def create_placeholder(image_path, width=100, height=100):
#     placeholder = Image.new('RGB', (width, height), color='lightgray')
#     draw = ImageDraw.Draw(placeholder)
#     draw.text((width//4, height//3), "No Image", fill="black")
#     placeholder.save(image_path)

# # Load your workbook and sheet
# wb = load_workbook('Tools List-2.xlsx')
# sheet = wb['Sheet1']

# # Create a SheetImageLoader instance
# image_loader = SheetImageLoader(sheet)

# # Create an 'images' directory if it doesn't exist
# if not os.path.exists('images'):
#     os.makedirs('images')

# # Loop through the C column (C1, C2, C3, etc.)
# for row in range(1, sheet.max_row + 1):
#     cell = f'C{row}'
#     try:
#         # Try to get the image in the specified cell
#         image = image_loader.get(cell)

#         if image:
#             # Save the image if found
#             image_filename = f'images/image_{row}.png'
#             image.save(image_filename)
#             print(f"Image saved for {cell}")
#         else:
#             # If no image, create and save a placeholder image
#             image_filename = f'images/image_{row}.png'
#             create_placeholder(image_filename)
#             print(f"Placeholder saved for {cell}")
#     except Exception as e:
#         # Handle any exception (if any issue arises during image extraction)
#         print(f"Error processing {cell}: {e}")

# shutil.rmtree("images", ignore_errors=True)


# '''final'''




# import os
# import io
# import shutil
# import pandas as pd
# from PIL import Image, ImageDraw
# from openpyxl import load_workbook
# from openpyxl_image_loader import SheetImageLoader
# from reportlab.lib import colors
# from reportlab.lib.pagesizes import A4
# from reportlab.pdfgen import canvas
# import textwrap

# # Page and grid settings
# PAGE_WIDTH, PAGE_HEIGHT = A4
# MARGIN = 30
# GRID_COLS = 3
# GRID_ROWS = 4
# CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
# CELL_HEIGHT = (PAGE_HEIGHT - 120) / GRID_ROWS  # Adjusted for header
# IMAGE_HEIGHT = CELL_HEIGHT * 0.6
# IMAGE_WIDTH = CELL_WIDTH - 20

# # Fonts and colors
# font_heading = "Helvetica-Bold"
# font_text = "Helvetica"
# font_size_heading = 24
# font_size_address = 10
# font_size_name = 10
# font_size_price = 12
# color_heading = colors.HexColor("#2C3E50")
# color_text = colors.HexColor("#34495E")
# color_price = colors.HexColor("#27AE60")
# bg_color_cell = colors.HexColor("#ECF0F1")
# border_color_cell = colors.HexColor("#BDC3C7")

# # Function to wrap text
# def wrap_text(text, width):
#     return textwrap.fill(text, width)

# # Function to draw the grid with enhanced styling
# def draw_grid(canvas_obj, start_y):
#     for row in range(GRID_ROWS):
#         for col in range(GRID_COLS):
#             x = MARGIN + col * CELL_WIDTH
#             y = start_y - row * CELL_HEIGHT
#             # Draw cell background
#             canvas_obj.setFillColor(bg_color_cell)
#             canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=True, stroke=False)
#             # Draw cell border
#             canvas_obj.setStrokeColor(border_color_cell)
#             canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=False, stroke=True)

# # Function to create the PDF
# def create_pdf(data, placeholder_image, output_pdf):
#     c = canvas.Canvas(output_pdf, pagesize=A4)
#     items_per_page = GRID_ROWS * GRID_COLS
#     total_items = len(data)

#     for page_num in range((total_items + items_per_page - 1) // items_per_page):
#         if page_num > 0:
#             c.showPage()

#         # Add heading
#         c.setFont(font_heading, font_size_heading)
#         c.setFillColor(color_heading)
#         c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")

#         # Add address
#         c.setFont(font_text, font_size_address)
#         c.setFillColor(color_text)
#         c.drawCentredString(
#             PAGE_WIDTH / 2,
#             PAGE_HEIGHT - 70,
#             "Address: 123 Example Street, City Name, State, Country"
#         )

#         # Draw grid
#         grid_start_y = PAGE_HEIGHT - 100
#         draw_grid(c, grid_start_y)

#         # Process items for the current page
#         start_idx = page_num * items_per_page
#         end_idx = min((page_num + 1) * items_per_page, total_items)

#         for idx in range(start_idx, end_idx):
#             row = data.iloc[idx]
#             current_row = (idx - start_idx) // GRID_COLS
#             current_col = (idx - start_idx) % GRID_COLS

#             # Get product details
#             product_name = str(row.get("Name", "Unknown Product"))
#             price = str(row.get("Price E", "Unknown Price"))
#             image_name = row.get("Image Name", None)

#             # Calculate cell position
#             x = MARGIN + current_col * CELL_WIDTH
#             y = grid_start_y - current_row * CELL_HEIGHT

#             # Draw image
#             try:
#                 # Construct the full image path
#                 image_path = os.path.join("images", image_name) if image_name else placeholder_image

#                 # Check if the image exists
#                 if os.path.exists(image_path):
#                     img = Image.open(image_path)
#                 else:
#                     print(f"Image not found: {image_path}, using placeholder.")
#                     img = Image.open(placeholder_image)

#                 # Convert and resize the image
#                 img = img.convert("RGB")
#                 img.resize((int(IMAGE_WIDTH), int(IMAGE_HEIGHT)), Image.Resampling.LANCZOS)

#                 # Save to a temporary file
#                 temp_image_path = f"temp_image_{idx}.jpg"
#                 img.save(temp_image_path)

#                 # Draw the image from the temporary file path
#                 c.drawImage(
#                     temp_image_path,
#                     x + (CELL_WIDTH - IMAGE_WIDTH) / 2,
#                     y - IMAGE_HEIGHT - 10,
#                     width=IMAGE_WIDTH,
#                     height=IMAGE_HEIGHT,
#                 )

#                 # Cleanup the temporary file after usage
#                 os.remove(temp_image_path)
#             except Exception as e:
#                 print(f"Error loading image for {product_name}: {e}")

#             # Draw product name
#             c.setFont(font_text, font_size_name)
#             c.setFillColor(color_text)
#             wrapped_name = wrap_text(product_name, 25)
#             text_y = y - IMAGE_HEIGHT - 20
#             for line in wrapped_name.splitlines():
#                 c.drawString(x + 10, text_y, line)
#                 text_y -= 12

#             # Draw price
#             c.setFont(font_text, font_size_price)
#             c.setFillColor(color_price)
#             c.drawString(x + 10, text_y - 5, f"₹ {price}")

#     c.save()
#     print(f"PDF generated successfully: {output_pdf}")

# # Extract images using updated method
# def extract_images_and_add_column(excel_file, output_excel_file):
#     wb = load_workbook(excel_file)
#     sheet = wb['Sheet1']
#     image_loader = SheetImageLoader(sheet)
#     image_dir = "images"

#     if not os.path.exists(image_dir):
#         os.makedirs(image_dir)

#     # Extract images from column C and save them
#     image_column = []
#     for row in range(2, sheet.max_row + 1):
#         cell = f'C{row}'
#         try:
#             image = image_loader.get(cell)
#             if image:
#                 image_filename = f"image_{row}.png"
#                 image_path = os.path.join(image_dir, image_filename)
#                 image.save(image_path)
#                 image_column.append(image_filename)
#             else:
#                 image_column.append("")  # No image
#         except Exception as e:
#             print(f"Error processing cell {cell}: {e}")
#             image_column.append("")  # No image

#     # Convert sheet to DataFrame
#     data = pd.DataFrame(sheet.values)
#     headers = data.iloc[0]
#     data = data[1:]
#     data.columns = headers

#     # Add image column
#     data["Image Name"] = image_column
#     data.to_excel(output_excel_file, index=False)
#     print(f"Updated Excel file saved with images in '{output_excel_file}'.")

# # Main Script
# excel_file = "Tools List-2.xlsx"
# updated_excel_file = "updated_test.xlsx"
# placeholder_image_path = "placeholder.jpg"
# output_pdf_path = "stylish_layout_output.pdf"

# # Create a placeholder image
# placeholder = Image.new("RGB", (300, 300), color="lightgray")
# draw = ImageDraw.Draw(placeholder)
# draw.text((100, 140), "Image", fill="black")
# placeholder.save(placeholder_image_path)

# # Extract images and add the "Image Name" column
# # extract_images_and_add_column(excel_file, updated_excel_file)

# # Generate the PDF
# data = pd.read_excel(updated_excel_file)
# create_pdf(data, placeholder_image_path, output_pdf_path)


# shutil.rmtree("images", ignore_errors=True)






# '''app and ui'''


import streamlit as st
import os
import io
import shutil
import pandas as pd
from PIL import Image, ImageDraw
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import textwrap

# Page and grid settings
PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 30
GRID_COLS = 3
GRID_ROWS = 4
CELL_WIDTH = (PAGE_WIDTH - 2 * MARGIN) / GRID_COLS
CELL_HEIGHT = (PAGE_HEIGHT - 120) / GRID_ROWS
IMAGE_HEIGHT = CELL_HEIGHT * 0.6
IMAGE_WIDTH = CELL_WIDTH - 20

# Fonts and colors
font_heading = "Helvetica-Bold"
font_text = "Helvetica"
font_size_heading = 24
font_size_address = 10
font_size_name = 10
font_size_price = 12
color_heading = colors.HexColor("#2C3E50")
color_text = colors.HexColor("#34495E")
color_price = colors.HexColor("#27AE60")
bg_color_cell = colors.HexColor("#ECF0F1")
border_color_cell = colors.HexColor("#BDC3C7")

def wrap_text(text, width):
    return textwrap.fill(text, width)

def draw_grid(canvas_obj, start_y):
    for row in range(GRID_ROWS):
        for col in range(GRID_COLS):
            x = MARGIN + col * CELL_WIDTH
            y = start_y - row * CELL_HEIGHT
            canvas_obj.setFillColor(bg_color_cell)
            canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=True, stroke=False)
            canvas_obj.setStrokeColor(border_color_cell)
            canvas_obj.rect(x, y - CELL_HEIGHT, CELL_WIDTH, CELL_HEIGHT, fill=False, stroke=True)

def create_pdf(data, placeholder_image, output_pdf):
    c = canvas.Canvas(output_pdf, pagesize=A4)
    items_per_page = GRID_ROWS * GRID_COLS
    total_items = len(data)

    for page_num in range((total_items + items_per_page - 1) // items_per_page):
        if page_num > 0:
            c.showPage()

        c.setFont(font_heading, font_size_heading)
        c.setFillColor(color_heading)
        c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 50, "MobileClinic")

        c.setFont(font_text, font_size_address)
        c.setFillColor(color_text)
        c.drawCentredString(
            PAGE_WIDTH / 2,
            PAGE_HEIGHT - 70,
            "Address: 123 Example Street, City Name, State, Country"
        )

        grid_start_y = PAGE_HEIGHT - 100
        draw_grid(c, grid_start_y)

        start_idx = page_num * items_per_page
        end_idx = min((page_num + 1) * items_per_page, total_items)

        for idx in range(start_idx, end_idx):
            row = data.iloc[idx]
            current_row = (idx - start_idx) // GRID_COLS
            current_col = (idx - start_idx) % GRID_COLS

            product_name = str(row.get("Name", "Unknown Product"))
            price = str(row.get("Price E", "Unknown Price"))
            image_name = row.get("Image Name", None)

            x = MARGIN + current_col * CELL_WIDTH
            y = grid_start_y - current_row * CELL_HEIGHT

            try:
                image_path = os.path.join("images", image_name) if image_name else placeholder_image

                if os.path.exists(image_path):
                    img = Image.open(image_path)
                else:
                    img = Image.open(placeholder_image)

                img = img.convert("RGB")
                img.resize((int(IMAGE_WIDTH), int(IMAGE_HEIGHT)), Image.Resampling.LANCZOS)

                temp_image_path = f"temp_image_{idx}.jpg"
                img.save(temp_image_path)

                c.drawImage(
                    temp_image_path,
                    x + (CELL_WIDTH - IMAGE_WIDTH) / 2,
                    y - IMAGE_HEIGHT - 10,
                    width=IMAGE_WIDTH,
                    height=IMAGE_HEIGHT,
                )

                os.remove(temp_image_path)
            except Exception as e:
                st.error(f"Error loading image for {product_name}: {e}")

            c.setFont(font_text, font_size_name)
            c.setFillColor(color_text)
            wrapped_name = wrap_text(product_name, 25)
            text_y = y - IMAGE_HEIGHT - 20
            for line in wrapped_name.splitlines():
                c.drawString(x + 10, text_y, line)
                text_y -= 12

            c.setFont(font_text, font_size_price)
            c.setFillColor(color_price)
            c.drawString(x + 10, text_y - 5, f"₹ {price}")

    c.save()
    return output_pdf

def extract_images_and_add_column(excel_file):
    wb = load_workbook(excel_file)
    sheet = wb['Sheet1']
    image_loader = SheetImageLoader(sheet)
    image_dir = "images"

    if not os.path.exists(image_dir):
        os.makedirs(image_dir)

    image_column = []
    for row in range(2, sheet.max_row + 1):
        cell = f'C{row}'
        try:
            image = image_loader.get(cell)
            if image:
                image_filename = f"image_{row}.png"
                image_path = os.path.join(image_dir, image_filename)
                image.save(image_path)
                image_column.append(image_filename)
            else:
                image_column.append("")
        except Exception as e:
            st.error(f"Error processing cell {cell}: {e}")
            image_column.append("")

    data = pd.DataFrame(sheet.values)
    headers = data.iloc[0]
    data = data[1:]
    data.columns = headers
    data["Image Name"] = image_column
    
    return data

def main():
    st.title("MobileClinic Product Catalog Generator")
    
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
    
    if uploaded_file:
        # Create placeholder image
        placeholder = Image.new("RGB", (300, 300), color="lightgray")
        draw = ImageDraw.Draw(placeholder)
        draw.text((100, 140), "Image", fill="black")
        placeholder_path = "placeholder.jpg"
        placeholder.save(placeholder_path)
        
        if st.button("Generate PDF"):
            with st.spinner("Processing..."):
                # Save uploaded file temporarily
                with open("temp.xlsx", "wb") as f:
                    f.write(uploaded_file.getvalue())
                
                # Process the Excel file
                data = extract_images_and_add_column("temp.xlsx")
                
                # Generate PDF
                output_pdf = "catalog.pdf"
                create_pdf(data, placeholder_path, output_pdf)
                
                # Read PDF for download
                with open(output_pdf, "rb") as f:
                    pdf_bytes = f.read()
                
                # Clean up
                os.remove("temp.xlsx")
                os.remove(placeholder_path)
                os.remove(output_pdf)
                shutil.rmtree("images", ignore_errors=True)
                
                # Offer download
                st.success("PDF generated successfully!")
                st.download_button(
                    label="Download PDF",
                    data=pdf_bytes,
                    file_name="catalog.pdf",
                    mime="application/pdf"
                )

if __name__ == "__main__":
    main()