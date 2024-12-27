import cv2
import numpy as np

def open_pdf_viewer():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            image = first_page.to_image(resolution=150).original

            # Convert image to numpy array
            img_np = np.array(image)

            # Define lower and upper bounds for red color in RGB format
            lower_red = np.array([0, 0, 200])
            upper_red = np.array([50, 50, 255])

            # Create a binary mask for red color
            mask = cv2.inRange(img_np, lower_red, upper_red)

            # Find contours of red boxes
            contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

            # Extracted text from red boxes
            extracted_text = ""

            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                cropped_image = image.crop((x, y, x + w, y + h))
                text = pytesseract.image_to_string(cropped_image, config='--psm 6')  # You need to install pytesseract and Tesseract OCR
                extracted_text += text + "\n"

            text_dialog = tk.Toplevel(root)
            text_dialog.title("Selected Text from PDF")
            text_widget = tk.Text(text_dialog)
            text_widget.insert(tk.END, extracted_text)
            text_widget.pack()
