import tkinter as tk
from tkinter import filedialog
from docx import Document
import os
from PIL import Image
import shutil

def extract_images(document, output_folder):
    """Extract images from the Word document and save them to the output folder."""
    rels = document.part.rels
    image_index = 1
    image_paths = []

    for rel in rels.values():
        if "image" in rel.target_ref:
            image_path = os.path.join(output_folder, f"image_{image_index}.png")
            image_index += 1
            with open(image_path, "wb") as img:
                img.write(rel.target_part.blob)
            image_paths.append(image_path)

    return image_paths

def word_to_markdown(doc_path, output_folder):
    """Convert the Word document to Markdown format."""
    # Load the document
    document = Document(doc_path)

    # Create an output folder for images
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Extract images
    image_paths = extract_images(document, output_folder)

    # Generate Markdown content
    markdown_content = ""

    # Add headers and paragraphs
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            markdown_content += text + "\n\n"

    # Include images in Markdown
    for idx, image_path in enumerate(image_paths):
        markdown_content += f"![Image {idx + 1}]({image_path})\n\n"

    return markdown_content

def save_markdown(content, output_file):
    """Save the Markdown content to a file."""
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(content)

def browse_file():
    """Handle file browsing and conversion."""
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if not file_path:
        return

    # Define output paths
    output_folder = "markdown_output"
    markdown_file = os.path.join(output_folder, "output.md")

    # Convert to Markdown
    markdown_content = word_to_markdown(file_path, output_folder)

    # Save Markdown
    save_markdown(markdown_content, markdown_file)

    # Notify user
    tk.messagebox.showinfo("Success", f"Markdown file saved at {markdown_file}")

# Create Tkinter app
app = tk.Tk()
app.title("Word to Markdown Converter")
app.geometry("400x200")

# Add button to browse files
browse_button = tk.Button(app, text="Select Word File", command=browse_file)
browse_button.pack(pady=50)

app.mainloop()
