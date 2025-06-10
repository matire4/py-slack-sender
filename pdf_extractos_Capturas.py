import fitz  # PyMuPDF
from PIL import Image
import re
import os

# --- Configuration Module ---
class Config:
    """
    Configuration settings for the PDF processing.
    """
    OUTPUT_DIR = "output_captures"
    
    # CAMBIO 1 (PREVENTIVO): Usamos la expresión regular mejorada para evitar capturar texto extra.
    PERSON_HEADER_PATTERN = re.compile(r"Consumos\s+(.*?)(?:\s{2,}|FECHA|DESCRIPCIÓN|NRO\. CUPÓN|Banco|$)")

    TOTAL_CONSUMOS_PATTERN = "TOTAL CONSUMOS DE"
    START_MARKER = "DETALLE"
    END_MARKER_1 = "Impuestos, cargos e intereses"
    END_MARKER_2 = "Legales y avisos"
    DEFAULT_FILENAME_PLACEHOLDER = "Persona XXX"
    DPI = 300 # Higher DPI for better image quality

# --- PDF Processing Module ---
class PDFProcessor:
    """
    Handles reading and processing the PDF to identify relevant sections.
    """
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.document = fitz.open(pdf_path)
        self.relevant_page_range = self._get_relevant_page_range()

    def _get_relevant_page_range(self):
        """
        Identifies the start page for processing. The end is handled
        dynamically within find_person_sections based on end markers.
        """
        start_idx = -1
        # Find start page
        for i in range(self.document.page_count):
            page = self.document.load_page(i)
            if Config.START_MARKER in page.get_text():
                start_idx = i
                break

        if start_idx == -1:
            raise ValueError(f"'{Config.START_MARKER}' marker not found in the PDF. Cannot determine start page.")

        # Set end_idx to the full document count.
        # Dynamic termination will happen in find_person_sections.
        end_idx = self.document.page_count 

        return start_idx, end_idx

    def find_person_sections(self):
        sections = []
        current_section = None
        start_page_idx, _ = self.relevant_page_range # end_page_idx is ignored, loop goes to end
        
        stop_processing_consumption = False # NEW FLAG

        for i in range(start_page_idx, self.document.page_count): # Loop to the very end of the document
            page = self.document.load_page(i)
            words = page.get_text("words")
            lines_with_bboxes = self._get_lines_with_bboxes(words)

            for line_text, line_bbox in lines_with_bboxes:
                # If we've hit an end marker previously, or on this line, stop all consumption processing
                if stop_processing_consumption:
                    break # Break from inner line loop

                # Check for end markers on the current line
                is_end_marker_line = (Config.END_MARKER_1 in line_text or Config.END_MARKER_2 in line_text)
                
                # If an end marker line is encountered
                if is_end_marker_line:
                    stop_processing_consumption = True # Set flag to stop further consumption processing

                    # Finalize any open section before stopping
                    if current_section and current_section['end_bbox'] is None:
                        if current_section['details_bboxes']:
                            _, last_detail_bbox = current_section['details_bboxes'][-1]
                            current_section['end_page'] = current_section['details_bboxes'][-1][0]
                            current_section['end_bbox'] = last_detail_bbox
                        else:
                            current_section['end_page'] = current_section['start_page']
                            current_section['end_bbox'] = current_section['start_bbox'] # Fallback
                        print(f"WARNING: Section for '{current_section['name']}' started on page {current_section['start_page']+1} ended due to '{line_text.strip()}' marker. Capturing up to here.")
                        sections.append(current_section)
                        current_section = None # No more consumption sections after this
                    break # Break from inner line loop, stop processing this page for consumption data

                # Process consumption data if we haven't stopped yet
                person_match = Config.PERSON_HEADER_PATTERN.search(line_text)
                total_match_start_idx = line_text.find(Config.TOTAL_CONSUMOS_PATTERN)

                if person_match:
                    # New person section found, finalize previous one if exists
                    if current_section:
                        if current_section['end_bbox'] is None:
                            # If a section was open but no total was found, it's an error/incomplete section
                            if current_section['details_bboxes']:
                                _, last_detail_bbox = current_section['details_bboxes'][-1]
                                current_section['end_page'] = current_section['details_bboxes'][-1][0]
                                current_section['end_bbox'] = last_detail_bbox
                            else:
                                current_section['end_page'] = current_section['start_page']
                                current_section['end_bbox'] = current_section['start_bbox'] # Use header bbox as fallback
                            print(f"WARNING: Section for '{current_section['name']}' started on page {current_section['start_page']+1} ended abruptly without 'TOTAL' line. Capturing up to here.")
                        sections.append(current_section)
                    
                    # CAMBIO 2 (LA SOLUCIÓN PRINCIPAL): Convertimos el nombre a mayúsculas
                    # para que coincida con la línea TOTAL, que también está en mayúsculas.
                    person_name = person_match.group(1).strip().upper()

                    current_section = {
                        'name': person_name,
                        'start_page': i,
                        'start_bbox': fitz.Rect(line_bbox),
                        'details_bboxes': [],
                        'end_page': None,
                        'end_bbox': None
                    }
                    continue # Move to next line, this line is a header

                elif current_section: # Only process as detail if a section is active
                    # Accumulate detail lines
                    if total_match_start_idx != -1 and line_text.strip().startswith(f"{Config.TOTAL_CONSUMOS_PATTERN} {current_section['name']}"):
                        # Found total line for current person
                        current_section['end_page'] = i
                        current_section['end_bbox'] = fitz.Rect(line_bbox)
                        sections.append(current_section)
                        current_section = None # Reset for next person
                    else:
                        # Add as a detail line
                        current_section['details_bboxes'].append((i, fitz.Rect(line_bbox)))
            
            # After processing all lines on a page, if stop_processing_consumption is true,
            # break from the outer page loop too.
            if stop_processing_consumption:
                break

        # Final check for any remaining open section after the loop (should be handled by `is_end_marker_line` now)
        if current_section:
            if current_section['end_bbox'] is None:
                if current_section['details_bboxes']:
                    last_detail_page, last_detail_bbox = current_section['details_bboxes'][-1]
                    current_section['end_page'] = last_detail_page
                    current_section['end_bbox'] = last_detail_bbox
                else:
                    current_section['end_page'] = current_section['start_page']
                    current_section['end_bbox'] = current_section['start_bbox']
                print(f"WARNING: Section for '{current_section['name']}' started on page {current_section['start_page']+1} ended unexpectedly at document end. Capturing up to here.")
            sections.append(current_section)

        # Post-processing for sections where the 'TOTAL' line was not found:
        # This block primarily catches edge cases where the end_bbox might still be None
        # despite previous logic.
        for section in sections:
            if section['end_bbox'] is None:
                print(f"WARNING: 'Alguno de los detalles de los vendedores tiene otro formato, revisar el archivo' for {section['name']} (page {section['start_page']+1}). Image boundaries might be approximate.")
                if section['details_bboxes']:
                    last_detail_page, last_detail_bbox = section['details_bboxes'][-1]
                    section['end_page'] = last_detail_page
                    section['end_bbox'] = last_detail_bbox
                else:
                    section['end_page'] = section['start_page']
                    section['end_bbox'] = section['start_bbox']
        return sections

    def _get_lines_with_bboxes(self, words):
        """
        Groups words into lines and associates a bounding box with each line.
        """
        lines = {}
        for w in words:
            # w = (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            key = (w[5], w[6]) # block_no, line_no
            if key not in lines:
                lines[key] = {'text': [], 'bbox': fitz.Rect(w[0], w[1], w[2], w[3])}
            lines[key]['text'].append(w[4])
            lines[key]['bbox'].include_rect(fitz.Rect(w[0], w[1], w[2], w[3]))

        sorted_lines = sorted(lines.items(), key=lambda x: (x[0][0], x[0][1]))
        return [(" ".join(item[1]['text']), item[1]['bbox']) for item in sorted_lines]

    def close(self):
        self.document.close()

# --- Image Generation Module ---
class ImageGenerator:
    """
    Responsible for rendering PDF sections into cropped images.
    """
    def __init__(self, document):
        self.document = document
        os.makedirs(Config.OUTPUT_DIR, exist_ok=True)
        self.name_counts = {} # New: To track occurrences of names

    def generate_image(self, section_data):
        """
        Generates and saves a cropped image for a given person's section.
        """
        person_name = section_data['name']
        start_page_idx = section_data['start_page']
        end_page_idx = section_data['end_page']
        start_bbox = section_data['start_bbox']
        end_bbox = section_data['end_bbox']
        details_bboxes = section_data['details_bboxes']

        # Determine base filename
        base_filename = f"Consumos {person_name}" if person_name else Config.DEFAULT_FILENAME_PLACEHOLDER

        # Handle duplicate names for unique file paths
        if base_filename in self.name_counts:
            self.name_counts[base_filename] += 1
            filename = f"{base_filename} ({self.name_counts[base_filename]}).jpg"
        else:
            self.name_counts[base_filename] = 1
            filename = f"{base_filename}.jpg"

        output_path = os.path.join(Config.OUTPUT_DIR, filename)

        if start_page_idx == end_page_idx:
            # Single page section
            page = self.document.load_page(start_page_idx)
            pix = page.get_pixmap(matrix=fitz.Matrix(Config.DPI/72, Config.DPI/72))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Calculate the overall bounding box for cropping on a single page
            # Ahora que la agrupación de secciones es correcta, esta lógica funcionará bien.
            overall_top = start_bbox.y0
            overall_bottom = end_bbox.y1
            overall_left = min(start_bbox.x0, end_bbox.x0)
            overall_right = max(start_bbox.x1, end_bbox.x1)

            for p_idx, d_bbox in details_bboxes:
                if p_idx == start_page_idx: # Ensure it's on the current page
                    overall_top = min(overall_top, d_bbox.y0)
                    overall_bottom = max(overall_bottom, d_bbox.y1)
                    overall_left = min(overall_left, d_bbox.x0)
                    overall_right = max(overall_right, d_bbox.x1)
            
            # Convert PDF coordinates to image pixel coordinates
            scale_factor = Config.DPI / 72
            crop_box = (
                int(overall_left * scale_factor)-30,
                int(overall_top * scale_factor)-30,
                int(overall_right * scale_factor)+30,
                int(overall_bottom * scale_factor)+30
            )
            img_cropped = img.crop(crop_box)
            img_cropped.save(output_path, dpi=(Config.DPI, Config.DPI))
            print(f"Generated: {output_path}")

        else:
            # Multi-page section - requires stitching
            images_to_stitch = []
            
            # First page segment
            page_start = self.document.load_page(start_page_idx)
            pix_start = page_start.get_pixmap(matrix=fitz.Matrix(Config.DPI/72, Config.DPI/72))
            img_start = Image.frombytes("RGB", [pix_start.width, pix_start.height], pix_start.samples)
            
            # Calculate crop box for the first page: from start_bbox.y0 to bottom of page
            scale_factor = Config.DPI / 72
            crop_start_page = (
                int(page_start.rect.x0 * scale_factor),
                int(start_bbox.y0 * scale_factor),
                int(page_start.rect.x1 * scale_factor),
                int(page_start.rect.y1 * scale_factor)
            )
            images_to_stitch.append(img_start.crop(crop_start_page))

            # Intermediate pages
            for p_idx in range(start_page_idx + 1, end_page_idx):
                page_inter = self.document.load_page(p_idx)
                pix_inter = page_inter.get_pixmap(matrix=fitz.Matrix(Config.DPI/72, Config.DPI/72))
                images_to_stitch.append(Image.frombytes("RGB", [pix_inter.width, pix_inter.height], pix_inter.samples))

            # Last page segment
            page_end = self.document.load_page(end_page_idx)
            pix_end = page_end.get_pixmap(matrix=fitz.Matrix(Config.DPI/72, Config.DPI/72))
            img_end = Image.frombytes("RGB", [pix_end.width, pix_end.height], pix_end.samples)

            # Calculate crop box for the last page: from top of page to end_bbox.y1
            crop_end_page = (
                int(page_end.rect.x0 * scale_factor),
                int(page_end.rect.y0 * scale_factor),
                int(page_end.rect.x1 * scale_factor),
                int(end_bbox.y1 * scale_factor)
            )
            images_to_stitch.append(img_end.crop(crop_end_page))

            # Stitch images vertically
            widths, heights = zip(*(i.size for i in images_to_stitch))
            max_width = max(widths)
            total_height = sum(heights)

            stitched_image = Image.new('RGB', (max_width, total_height))
            y_offset = 0
            for img in images_to_stitch:
                stitched_image.paste(img, (0, y_offset))
                y_offset += img.size[1]
            
            # Further refinement: the stitching above creates an image that is composed of segments.
            # We want to ensure the final image is exactly the width of the content.
            # We can re-calculate the overall left/right from the bounding boxes to refine the width.
            overall_left = min(start_bbox.x0, end_bbox.x0)
            overall_right = max(start_bbox.x1, end_bbox.x1)
            for p_idx, d_bbox in details_bboxes:
                overall_left = min(overall_left, d_bbox.x0)
                overall_right = max(overall_right, d_bbox.x1)
            
            # Adjust the stitched image's width based on content
            crop_stitched_width = (
                int(overall_left * scale_factor) - 30,
                0, # From top of stitched image
                int(overall_right * scale_factor) + 30,
                total_height # To bottom of stitched image
            )
            final_cropped_image = stitched_image.crop(crop_stitched_width)

            final_cropped_image.save(output_path, dpi=(Config.DPI, Config.DPI))
            print(f"Generated (stitched): {output_path}")


# --- Main Application Logic ---
def main(pdf_file_path):
    """
    Main function to orchestrate the PDF processing and image generation.
    """

    if not os.path.exists(pdf_file_path):
        print(f"Error: PDF file not found at '{pdf_file_path}'")
        return

    processor = None
    try:
        # Use a dynamic path for the PDF
        processor = PDFProcessor(pdf_file_path)
        person_sections = processor.find_person_sections()

        if not person_sections:
            print("No consumption sections found in the PDF.")
            return

        image_gen = ImageGenerator(processor.document)
        for section in person_sections:
            image_gen.generate_image(section)

        print(f"\nPDF processing complete. Images saved in '{Config.OUTPUT_DIR}' directory.")

    except ValueError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        if processor:
            processor.close()

# --- Script Execution ---
if __name__ == "__main__":
    # Asegúrate de que el nombre del archivo PDF aquí sea el correcto.
    pdf_path = '04-2025 - Gastos.pdf'
    main(pdf_path)