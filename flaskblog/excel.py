from PIL import ImageGrab
import xlwings as xw

def excel_catch_screen(shot_excel, shot_sheetname):
    app = xw.App(visible=True, add_book=False)          # Use xlwings Of app start-up
    wb = app.books.open(shot_excel)                     # Open file
    sheet = wb.sheets(shot_sheetname)                   # Selected sheet
    all = sheet.used_range                              # Get content range
    print(all.value)
    all.api.CopyPicture()                               # Copy picture area
    sheet.api.Paste()                                   # Paste
    img_name = 'data'
    pic = sheet.pictures[0]                             # Current picture
    pic.api.Copy()                                      # Copy the picture
    img = ImageGrab.grabclipboard()                     # Get the picture data of the clipboard
    img.save("C:\\Users\\Ynsnnn\\Desktop\\NewApp\\flaskblog\\Excel" + img_name + ".png")  # Save the picture
    pic.delete()                                        # Delete sheet Pictures on
    wb.close()                                          # Do not save , Direct closure
    app.quit()


excel_catch_screen("C:\\Users\\Ynsnnn\\Downloads\\test.xlsx", "Sheet1")