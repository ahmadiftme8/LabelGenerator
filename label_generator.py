import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import re
import logging
import arabic_reshaper
from bidi.algorithm import get_display

# همه شعب

class ShelfLabelGenerator:
    def __init__(self):
        # Base dimensions in pixels (at 300 DPI)
        self.WIDTH = 992  # 8.4 cm * 300 DPI/2.54 cm/inch
        self.HEIGHT = 508  # 4.3 cm * 300 DPI/2.54 cm/inch
        self.DPI = 300    # Set DPI for print quality
        
        # Text margins (at 300 DPI)
        self.TEXT_MARGIN = 83  # 0.7 cm * 300 DPI/2.54 cm/inch for main text margin
        self.INNER_TEXT_MARGIN = 71  # 0.6 cm * 300 DPI/2.54 cm/inch for price and product name
        
        # NEW margin for the product name specifically (0.3 cm)
        self.PRODUCT_NAME_LEFT_MARGIN = int(0.3 * 300 / 2.54)  # 0.3 cm * 300 DPI/2.54 cm/inch

        # Colors in RGB
        self.DARK_BG = (26, 26, 26)  # CSS: #1a1a1a
        self.GOLD_COLOR = (206, 175, 136)  # CSS: #ceaf88

        # Font colors
        self.PRICE_TAX_COLOR = (26, 26, 26)  # CSS: #1a1a1a
        self.PRODUCT_NAME_COLOR = (206, 175, 136)  # CSS: #ceaf88

        # Gold area dimensions (fixed, no margin)
        self.GOLD_AREA_WIDTH = 502  # Fixed width (47.6% of total width) #اصلش 472 بود

        self.GOLD_AREA_HEIGHT = 289  # Height remains the same

        # First initialize logging
        self.setup_logging()
        # Then setup directories
        self.setup_directories()
        # Finally load fonts
        self.load_fonts()

    def setup_logging(self):
        """Setup logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)

    def setup_directories(self):
        """Setup necessary directories"""
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.fonts_dir = os.path.join(self.current_dir, 'fonts')
        self.output_dir = os.path.join(self.current_dir, "generated_labels")
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.fonts_dir, exist_ok=True)

    def load_fonts(self):
        """Load required fonts"""
        try:
            font_path = os.path.join(self.fonts_dir, "PeydaFaNum-Bold.ttf")
            if not os.path.exists(font_path):
                raise FileNotFoundError(f"Font file not found: {font_path}")

            # Font sizes for 300 DPI
            self.fonts = {
                'product_name': ImageFont.truetype(font_path, 62), #درحالت عادی 58 است
                'price': ImageFont.truetype(font_path, 64),
                'plus': ImageFont.truetype(font_path, 60),
                'tax': ImageFont.truetype(font_path, 36),
            }
        except Exception as e:
            self.logger.error(f"Font loading error: {str(e)}")
            raise

    def create_label(self, product_name: str, price: float) -> Image.Image:
        try:
            # Create new image with the dark background
            image = Image.new('RGB', (self.WIDTH, self.HEIGHT), self.DARK_BG)
            draw = ImageDraw.Draw(image)

            # Draw gold area on the left (fixed position, no margin)
            gold_area = Image.new('RGB', (self.GOLD_AREA_WIDTH, self.GOLD_AREA_HEIGHT), self.GOLD_COLOR)
            image.paste(gold_area, (0, int((self.HEIGHT - self.GOLD_AREA_HEIGHT) / 2)))

            # Draw the price in the gold area (with left margin)
            price_text = f"{int(price):,} تومان"
            prepared_price_text = self.prepare_persian_text(price_text)
            # Calculate the available width for price text (gold area width minus left margin)
            price_area_width = self.GOLD_AREA_WIDTH - self.INNER_TEXT_MARGIN
            price_x = self.INNER_TEXT_MARGIN + (price_area_width // 2) - 24 
            self.draw_text_centered(draw, prepared_price_text, (price_x, 174), self.fonts['price'], self.PRICE_TAX_COLOR)

            # Product name area (starts after gold area plus product name left margin, ends before right margin)
            product_area_width = self.WIDTH - self.GOLD_AREA_WIDTH - self.PRODUCT_NAME_LEFT_MARGIN - self.INNER_TEXT_MARGIN - 10

            # Draw the product name with multiline support
            self.draw_multiline_text(
                draw, product_name,
                self.GOLD_AREA_WIDTH + self.PRODUCT_NAME_LEFT_MARGIN, (self.HEIGHT - 40) // 2,
                self.fonts['product_name'], self.PRODUCT_NAME_COLOR,
                product_area_width, self.HEIGHT - 40
            )

            # Draw plus sign and tax info (with left margin)
            prepared_tax_text = self.prepare_persian_text("%10 مالیات ارزش افزوده")
            plus_x = self.INNER_TEXT_MARGIN + (price_area_width // 2)
            self.draw_text_centered(draw, "+", (plus_x, 238), self.fonts['plus'], self.PRICE_TAX_COLOR)
            self.draw_text_centered(draw, prepared_tax_text, (plus_x, 312), self.fonts['tax'], self.PRICE_TAX_COLOR)

            # Set the DPI metadata
            image.info['dpi'] = (self.DPI, self.DPI)

            return image

        except Exception as e:
            self.logger.error(f"Label creation error: {str(e)}")
            return None

    def draw_multiline_text(self, draw: ImageDraw, text: str, x: int, y: int,
                        font: ImageFont, fill: tuple, max_width: int, max_height: int) -> int:
        """Draw text that wraps if it exceeds max width or has more than 3 words per line."""
        words = text.split()
        lines = []
        current_line = []

        for word in words:
            current_line.append(word)
            test_line = ' '.join(current_line)
            prepared_test = self.prepare_persian_text(test_line)
            width = draw.textlength(prepared_test, font=font)

            if width > max_width or len(current_line) > 3:
                if len(current_line) > 1:
                    current_line.pop()
                lines.append(' '.join(current_line))
                current_line = [word]

        if current_line:
            lines.append(' '.join(current_line))

        line_spacing = 24  # Increased for higher resolution
        total_height = 0
        line_heights = []

        for line in lines:
            prepared_line = self.prepare_persian_text(line)
            line_bbox = draw.textbbox((0, 0), prepared_line, font=font)
            line_height = line_bbox[3] - line_bbox[1]
            line_heights.append(line_height)
            total_height += line_height + line_spacing

        total_height -= line_spacing
        current_y = y - total_height // 2

        for i, line in enumerate(lines):
            prepared_line = self.prepare_persian_text(line)
            width = draw.textlength(prepared_line, font=font)
            x_centered = x + (max_width - width) // 2
            draw.text((x_centered, current_y), prepared_line, font=font, fill=fill)
            current_y += line_heights[i] + line_spacing

        return total_height

    def prepare_persian_text(self, text: str) -> str:
        """Prepare Persian text for correct rendering."""
        reshaped_text = arabic_reshaper.reshape(text)
        return get_display(reshaped_text)

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
                        label.save(filepath, quality=95, dpi=(self.DPI, self.DPI))
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
        """Clean filename by removing invalid characters and trimming whitespace."""
        logging.getLogger(__name__).warning(f"Cleaning filename for product: {filename}")
        filename = re.sub(r'[<>:"/\\|?*\t\n\r]', '', filename)  # Remove invalid characters and tabs/newlines
        return filename.strip().replace(' ', '_')  # Remove extra spaces and replace spaces with underscores


    # df = pd.read_excel("product_list.xlsx")
    # print(df.columns)
    # df.columns = df.columns.str.strip()
    # df = pd.read_excel("product_list.xlsx", encoding='utf-8')
   # df = pd.read_excel("product_list.xlsx", encoding='utf-8')

def main():
    try:
        generator = ShelfLabelGenerator()
        excel_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 
                                "product_list.xlsx")
        
        print("\n=== Label Generator Started ===\n")
        print("Requirements:")
        print("1. Persian font (PeydaFaNum-Bold.ttf) in 'fonts' folder")
        print("2. Excel file (product_list.xlsx) with columns:")
        print("   - 'نام محصول' (Product Name)")
        print("   - 'قیمت' (Price)")
        print(f"3. Output will be generated at {generator.DPI} DPI (8.4 cm x 4.3 cm)")
        
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


