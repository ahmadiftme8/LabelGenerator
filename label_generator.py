import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import re
import logging
import arabic_reshaper
from bidi.algorithm import get_display

class ShelfLabelGenerator:
    def __init__(self):
        # Fixed dimensions in pixels
        self.WIDTH = 992
        self.HEIGHT = 508

        # Colors in RGB (converted from CMYK)
        self.DARK_BG = (26, 26, 26)  # RGB equivalent of CMYK(0%, 0%, 0%, 95%)
        self.GOLD_COLOR = (184, 153, 113)  # RGB equivalent of CMYK(20%, 30%, 50%, 0%)

        # Font color for price and tax (black)
        self.PRICE_TAX_COLOR = (26, 26, 26)  # RGB equivalent of CMYK(0%, 0%, 0%, 95%)
        # Font color for product name (gold)
        self.PRODUCT_NAME_COLOR = (184, 153, 113)  # RGB equivalent of CMYK(20%, 30%, 50%, 0%)

        # Gold area dimensions (left side)
        self.GOLD_AREA_WIDTH = int(self.WIDTH * 0.4)  # 40% of total width
        self.GOLD_AREA_HEIGHT = self.HEIGHT

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
            font_path = os.path.join(self.fonts_dir, "Peyda-Bold.ttf")
            if not os.path.exists(font_path):
                raise FileNotFoundError(f"Font file not found: {font_path}")

            self.fonts = {
                'product_name': ImageFont.truetype(font_path, 72),
                'price': ImageFont.truetype(font_path, 64),
                'tax': ImageFont.truetype(font_path, 36)
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
            # Create new image in RGB mode
            image = Image.new('RGB', (self.WIDTH, self.HEIGHT), self.DARK_BG)
            draw = ImageDraw.Draw(image)

            # Draw gold area on the left
            gold_area = Image.new('RGB', (self.GOLD_AREA_WIDTH, self.GOLD_AREA_HEIGHT), self.GOLD_COLOR)
            image.paste(gold_area, (0, 0))

            # Calculate the width of the product name area (right side)
            product_area_width = self.WIDTH - self.GOLD_AREA_WIDTH - 40  # Padding on the right

            # Draw product name in the right area
            self.draw_multiline_text(
                draw, product_name,
                self.GOLD_AREA_WIDTH + 40, 100,  # Position it with some padding
                self.fonts['product_name'], self.PRODUCT_NAME_COLOR,
                product_area_width, self.HEIGHT - 40
            )

            # Draw price in the gold area
            price_text = f"{int(price):,} تومان"
            self.draw_multiline_text(
                draw, price_text,
                40, 40,  # Padding from top and left
                self.fonts['price'], self.PRICE_TAX_COLOR,
                self.GOLD_AREA_WIDTH, 200
            )

            # Draw tax information below the price in the gold area
            tax_text = "+\n۱۰٪ مالیات ارزش افزوده"
            self.draw_multiline_text(
                draw, tax_text,
                40, 200,  # Position under the price
                self.fonts['tax'], self.PRICE_TAX_COLOR,
                self.GOLD_AREA_WIDTH, 120
            )

            return image

        except Exception as e:
            self.logger.error(f"Label creation error: {str(e)}")
            return None

    def draw_multiline_text(self, draw: ImageDraw, text: str, x: int, y: int, 
                            font: ImageFont, fill: tuple, max_width: int, max_height: int) -> int:
        """Draw text that wraps if it exceeds max width."""
        words = text.split()
        lines = []
        current_line = []

        for word in words:
            current_line.append(word)
            test_line = ' '.join(current_line)
            prepared_test = self.prepare_persian_text(test_line)
            width = draw.textlength(prepared_test, font=font)

            if width > max_width:
                if len(current_line) > 1:
                    current_line.pop()
                lines.append(' '.join(current_line))
                current_line = [word]
            else:
                lines.append(test_line)
                current_line = []

        total_height = 0
        current_y = y

        for line in lines:
            prepared_line = self.prepare_persian_text(line)
            bbox = draw.textbbox((0, 0), prepared_line, font=font)
            line_height = bbox[3] - bbox[1]

            if total_height + line_height > max_height:
                break

            width = draw.textlength(prepared_line, font=font)
            x_centered = x + (max_width - width) // 2
            draw.text((x_centered, current_y), prepared_line, font=font, fill=fill)

            current_y += line_height
            total_height += line_height

        return total_height


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
        print("1. Persian font (Peyda-Bold.ttf) in 'fonts' folder")
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
