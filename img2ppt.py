import PySimpleGUI as sg
import win32com.client
from os.path import isfile, normpath
from tkinter import Tk


class Img2PPT():
    def __init__(self):
        # Get open PowerPoint Presentation. If none, open PowerPoint and add an empty slide
        try:
            self.ppt = win32com.client.GetActiveObject(
                "PowerPoint.Application").ActivePresentation
        except:
            self.ppt = win32com.client.Dispatch(
                "PowerPoint.Application").Presentations.Add()
            self._add_blank_slide()

        self.slide_width = self.ppt.Slides(1).Master.Width
        self.slide_height = self.ppt.Slides(1).Master.Height

        # Use tkinter for accessing the clipboard
        self.tk = Tk()
        self.tk.withdraw()

    def _add_blank_slide(self):
        """Adds a new slide at the end of the Presentation."""
        self.ppt.Slides.Add(self.get_slides_count() + 1, 12)

    def get_slides_count(self):
        """Returns the number of slides in the Presentation."""
        return self.ppt.Slides.Count

    def add_slide(self):
        self._add_blank_slide()

    def get_slides_amount_list(self):
        """Returns the range of number of slides as a list to be used in the Spin Element."""
        count = self.get_slides_count()
        if count == 1:
            return [count]
        elif count > 1:
            return list(range(1, count + 1))

    def paste_image(self, slide_number):
        """Adds a copied image to the target slide then clears clipboard."""
        try:
            clipboard = self.tk.clipboard_get()
            if isfile(clipboard) and clipboard.endswith('.jpg'):
                app.ppt.Slides(slide_number).Shapes.AddPicture(
                    normpath(clipboard), 0, 1, 0, 0)
                self.tk.clipboard_clear()
        except:
            pass

    def stretch_all(self, slide_number):
        """Stretches all images to fit in their Slide."""
        for shape in self.ppt.Slides(slide_number).Shapes:
            shape.LockAspectRatio = 0
            shape.width = self.slide_width
            shape.height = self.slide_height
            shape.left = 0
            shape.top = 0

    def fit_horizontal(self, slide_number):
        """Fits all images horizontally in their Slide."""
        shapes_count = self.ppt.Slides(slide_number).Shapes.Count
        previous_top = 0
        for shape in self.ppt.Slides(slide_number).Shapes:
            shape.LockAspectRatio = 0
            shape.width = self.slide_width
            shape.height = self.slide_height / shapes_count
            shape.left = 0
            shape.top = previous_top
            previous_top += shape.height

    def fit_vertical(self, slide_number):
        """Fits all images vertically in their Slide."""
        shapes_count = self.ppt.Slides(slide_number).Shapes.Count
        previous_left = 0
        for shape in self.ppt.Slides(slide_number).Shapes:
            shape.LockAspectRatio = 0
            shape.width = self.slide_width / shapes_count
            shape.height = self.slide_height
            shape.left = previous_left
            shape.top = 0
            previous_left += shape.width

    def fit_four(self, slide_number):
        """Ideally accepts four images and fits them in the four corners of the Slide."""
        shapes_count = self.ppt.Slides(slide_number).Shapes.Count / 2
        previous_left = 0
        previous_top = 0
        for shape in self.ppt.Slides(slide_number).Shapes:
            shape.LockAspectRatio = 0
            shape.width = self.slide_width / shapes_count
            shape.height = self.slide_height / shapes_count
            shape.left = previous_left
            shape.top = previous_top
            previous_left += shape.width
            if previous_left >= self.slide_width:
                previous_left = 0
                previous_top += shape.height

    def close(self):
        """Stops tkinter."""
        self.tk.destroy()


# Set UI theme and Element names
sg.set_options(font=('', 20))

sg.theme('DarkBlue')
slides_spin = 'slides_spin'
add_slide_button = 'Add Slide'
paste_img_button = 'Paste Image'
stretch_all_button = 'Stretch All'
fit_vertical_button = 'Fit Vertical'
fit_horizontal_button = 'Fit Horizontal'
fit_four_button = 'Fit Four'
exit_button = 'Exit'

# Instantiate Img2PPT and use some of its methods to populate the Spin Element
app = Img2PPT()

layout = [
    [sg.Spin(values=app.get_slides_amount_list(), initial_value=app.get_slides_count(),
             text_color='Black', key=slides_spin, enable_events=True, readonly=True)],
    [sg.HorizontalSeparator()],
    [sg.Button(add_slide_button), sg.Button(paste_img_button)],
    [sg.HorizontalSeparator()],
    [sg.Button(stretch_all_button), sg.Button(fit_vertical_button),
     sg.Button(fit_horizontal_button), sg.Button(fit_four_button)],
    [sg.HorizontalSeparator()],
    [sg.Button(exit_button)]
]

window = sg.Window('Img2PPT', layout)

# UI starts here
while True:
    event, values = window.read()

    # Makes sure that the Spin's values are updated before other operations
    window.Element(slides_spin).Update(values=app.get_slides_amount_list())
    slides_count = app.get_slides_count()
    # Change the value of the Spin to the current amount of slides
    # in case the user manually deleted some slides
    if int(window.Element(slides_spin).get()) > slides_count:
        window.Element(slides_spin).Update(value=slides_count)

    if event == sg.WIN_CLOSED or event == exit_button:
        app.close()
        break

    elif event == add_slide_button:
        app.add_slide()
        window.Element(slides_spin).Update(
            value=app.get_slides_count(), values=app.get_slides_amount_list())

    elif event == paste_img_button:
        app.paste_image(slide_number=window.Element(slides_spin).get())

    elif event == stretch_all_button:
        app.stretch_all(slide_number=window.Element(slides_spin).get())

    elif event == fit_vertical_button:
        app.fit_vertical(slide_number=window.Element(slides_spin).get())

    elif event == fit_horizontal_button:
        app.fit_horizontal(slide_number=window.Element(slides_spin).get())

    elif event == fit_four_button:
        app.fit_four(slide_number=window.Element(slides_spin).get())

window.close()
