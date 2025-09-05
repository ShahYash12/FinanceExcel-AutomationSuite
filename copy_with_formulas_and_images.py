
import xlwings as xw
import os

# Set folder paths
source_folder = "/Users/yashshah/Downloads/Shifting Project/Samples for Shifting Project"          # Folder with .xlsx files
template_path = "/Users/yashshah/Downloads/Shifting Project/Template for Shiting Project/Template.xlsm"      # .xlsm template with buttons/macros on first sheet
output_folder = "/Users/yashshah/Downloads/Shifting Project/Output For Shifting Project"        # Destination folder for filled .xlsm files

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

for file in os.listdir(source_folder):
    if file.endswith(".xlsx") and not file.startswith("~$"):
        source_file = os.path.join(source_folder, file)
        new_filename = os.path.splitext(file)[0] + "_filled.xlsm"
        new_file_path = os.path.join(output_folder, new_filename)

        app = xw.App(visible=False)
        try:
            # Open source file
            wb_source = app.books.open(source_file)

            # Open template
            wb_template = app.books.open(template_path)

            # Keep only the first sheet in template (with macros/buttons)
            while len(wb_template.sheets) > 1:
                wb_template.sheets[-1].delete()

            # Copy each sheet from source into the destination (preserve formulas/images)
            for sheet in wb_source.sheets:
                sheet.api.Copy(Before=wb_template.sheets[0].api)

            # Optional: rename the macro sheet if needed
            wb_template.sheets[-1].name = "Instructions"

            # Save as new file
            wb_template.save(new_file_path)
            print(f"✅ Created with full content: {new_filename}")

        except Exception as e:
            print(f"❌ Error processing {file}: {e}")
        finally:
            wb_source.close()
            wb_template.close()
            app.quit()
