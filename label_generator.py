import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import re
import logging
import arabic_reshaper
from bidi.algorithm import get_display

class ShelfLabelGenerator:
    def __init__(self):
        # Fixed dimensions in pixels (matching CSS)
        self.WIDTH = 992
        self.HEIGHT = 508

        # Colors in RGB (matching CSS)
        self.DARK_BG = (26, 26, 26)  # CSS: #1a1a1a
        self.GOLD_COLOR = (206, 175, 136)  # CSS: #ceaf88

        # Font color for price and tax (black)
        self.PRICE_TAX_COLOR = (26, 26, 26)  # CSS: #1a1a1a
        # Font color for product name (gold)
        self.PRODUCT_NAME_COLOR = (206, 175, 136)  # CSS: #ceaf88

        # Gold area dimensions (left side) (matching CSS)
        self.GOLD_AREA_WIDTH = 472
        self.GOLD_AREA_HEIGHT = 289

        # Setup logging and directories
        self._setup_logging()
        self._setup_directories()
        self._load_fonts()

    def _setup_logging(self):
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)

    def _setup_directories(self):
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.fonts_dir = os.path.join(self.current_dir, 'fonts')
        self.output_dir = os.path.join(self.current_dir, "generated_labels")
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.fonts_dir, exist_ok=True)

    def _load_fonts(self):
        try:
            font_path = os.path.join(self.fonts_dir, "PeydaFaNum-Bold.ttf")
            if not os.path.exists(font_path):
                raise FileNotFoundError(f"Font file not found: {font_path}")

            self.fonts = {
                'product_name': ImageFont.truetype(font_path, 60),  # Matching CSS: 70px
                'price': ImageFont.truetype(font_path, 60),  # Matching CSS: 60px
                'plus': ImageFont.truetype(font_path, 50),  # Matching CSS: 50px
                'tax': ImageFont.truetype(font_path, 30),  # Matching CSS: 30px
            }
        except Exception as e:
            self.logger.error(f"Font loading error: {str(e)}")
            raise

    def prepare_persian_text(self, text: str) -> str:
        """Prepare Persian text for correct rendering."""
        reshaped_text = arabic_reshaper.reshape(text)
        return get_display(reshaped_text)

    def create_label(self, product_name: str, price: float) -> Image.Image:
        try:
            # Create new image in RGB mode with the dark background
            image = Image.new('RGB', (self.WIDTH, self.HEIGHT), self.DARK_BG)
            draw = ImageDraw.Draw(image)

            # Draw gold area on the left
            gold_area = Image.new('RGB', (self.GOLD_AREA_WIDTH, self.GOLD_AREA_HEIGHT), self.GOLD_COLOR)
            image.paste(gold_area, (0, int((self.HEIGHT - self.GOLD_AREA_HEIGHT) / 2)))

            # Draw the price in the gold area (using reshaped text)
            price_text = f"{int(price):,} تومان"
            prepared_price_text = self.prepare_persian_text(price_text)  # Prepare price text
            self.draw_text_centered(draw, prepared_price_text, (236, 164), self.fonts['price'], self.PRICE_TAX_COLOR)

            # Define the width for the product name area
            product_area_width = self.WIDTH - self.GOLD_AREA_WIDTH - 40  # Update this line

            # Draw the product name with multiline support
            self.draw_multiline_text(
                draw, product_name,
                self.GOLD_AREA_WIDTH + 40, (self.HEIGHT - 40) // 2,  # Centered vertically
                self.fonts['product_name'], self.PRODUCT_NAME_COLOR,
                product_area_width, self.HEIGHT - 40
            )

            # Draw plus sign and tax info (using reshaped text)
            prepared_tax_text = self.prepare_persian_text("10% مالیات ارزش افزوده")
            self.draw_text_centered(draw, "+", (236, 224), self.fonts['plus'], self.PRICE_TAX_COLOR)
            self.draw_text_centered(draw, prepared_tax_text, (236, 294), self.fonts['tax'], self.PRICE_TAX_COLOR)

            return image

        except Exception as e:
            self.logger.error(f"Label creation error: {str(e)}")
        return None

    def draw_multiline_text(self, draw: ImageDraw, text: str, x: int, y: int,
                            font: ImageFont, fill: tuple, max_width: int, max_height: int) -> int:
        """Draw text that wraps if it exceeds max width or has more than 3 words per line."""

        words = text.split()  # Split the text into words
        lines = []
        current_line = []

        # Reduce the max_width by 50px for the right margin
        max_width -= 50

        for word in words:
            current_line.append(word)
            test_line = ' '.join(current_line)
            prepared_test = self.prepare_persian_text(test_line)
            width = draw.textlength(prepared_test, font=font)

            # Wrap the line if it's too wide or has more than 3 words
            if width > max_width or len(current_line) > 3:
                if len(current_line) > 1:
                    current_line.pop()  # Remove the last word to wrap it on the next line
                lines.append(' '.join(current_line))
                current_line = [word]

        # Add the last line
        if current_line:
            lines.append(' '.join(current_line))

        # Calculate total height of all lines
        total_height = 0
        line_heights = []

        for line in lines:
            prepared_line = self.prepare_persian_text(line)
            line_bbox = draw.textbbox((0, 0), prepared_line, font=font)
            line_height = line_bbox[3] - line_bbox[1]
            line_heights.append(line_height)
            total_height += line_height

        # Start drawing each line from the calculated y position
        current_y = y - total_height // 2  # Centering in Y-axis

        for i, line in enumerate(lines):
            prepared_line = self.prepare_persian_text(line)
            width = draw.textlength(prepared_line, font=font)
            x_centered = x + (max_width - width) // 2  # Centering in X-axis
            draw.text((x_centered, current_y), prepared_line, font=font, fill=fill)

            current_y += line_heights[i]  # Move to the next line's y position

        return total_height

    def draw_text_centered(self, draw: ImageDraw, text: str, position: tuple, font: ImageFont, fill: tuple):
        """Draw centered text at the specified position."""
        text_width = draw.textlength(text, font=font)
        x = position[0] - text_width // 2
        y = position[1] - (font.getbbox(text)[3] - font.getbbox(text)[1]) // 2
        draw.text((x, y), text, font=font, fill=fill)

    def generate_labels_from_excel(self, excel_path: str):
        """Generate labels from Excel file and save images."""
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        success_count = fail_count = 0

        try:
            df = pd.read_excel(excel_path)

            for index, row in df.iterrows():
                try:
                    product_name = str(row['نام محصول'])
                    price = float(re.sub(r'[^0-9.]', '', str(row['قیمت'])))

                    if label := self.create_label(product_name, price):
                        filename = f"{self.clean_filename(product_name)}.jpg"
                        filepath = os.path.join(self.output_dir, filename)
                        label.save(filepath, quality=95)
                        success_count += 1
                        self.logger.info(f"Generated: {filename}")
                    else:
                        fail_count += 1

                except Exception as e:
                    self.logger.error(f"Error processing row {index + 1}: {str(e)}")
                    fail_count += 1
                    continue

            self.logger.info(f"Generation complete: {success_count} successful, {fail_count} failed")
            return success_count, fail_count

        except Exception as e:
            self.logger.error(f"Excel processing error: {str(e)}")
            raise

    @staticmethod
    def clean_filename(filename: str) -> str:
        return re.sub(r'[<>:"/\\|?*]', '', filename).replace(' ', '_')

def main():
    try:
        generator = ShelfLabelGenerator()
        excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "product_list.xlsx")

        print("\n=== Label Generator Started ===\n")
        print("Requirements:")
        print("1. Persian font (PeydaFaNum-Bold.ttf) in 'fonts' folder")
        print("2. Excel file (product_list.xlsx) with columns:")
        print(" - 'نام محصول' (Product Name)")
        print(" - 'قیمت' (Price)")

        input("\nPress Enter to start...")

        success, failed = generator.generate_labels_from_excel(excel_file)
        print(f"\nGeneration complete!")
        print(f"Successfully generated: {success} labels")
        print(f"Failed: {failed} labels")
        print(f"Labels saved in: {generator.output_dir}")

    except Exception as e:
        print(f"\nError: {str(e)}")
    finally:
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
