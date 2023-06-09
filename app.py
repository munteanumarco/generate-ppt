import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton
from PyQt5.QtGui import QIcon
import subprocess

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # Create the input box and button
        self.url_input = QLineEdit(self)
        self.url_input.setGeometry(50, 50, 500, 30)
        self.url_input.setPlaceholderText("Enter URL")
        self.url_input.returnPressed.connect(self.generate_ppt)

        self.button = QPushButton('Generate PowerPoint', self)
        self.button.setGeometry(50, 100, 200, 30)
        
        self.setWindowTitle("Generate PPT")
        title_label = QLabel('Create PPT from Web Page', self)
        title_label.move(20, 10)
        title_label.resize(200, 20)

        script_dir = os.path.dirname(os.path.realpath(__file__))
        self.genppt_path = os.path.join(script_dir, 'genppt.py')

        self.req_path = os.path.join(script_dir, 'requirements.txt')
        #install dependencies 
        subprocess.run(['pip', 'install', '-r', self.req_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        # Connect the button to a function that runs the PowerPoint script
        self.button.clicked.connect(self.generate_ppt)

    def generate_ppt(self):
        # Get the URL from the input box
        url = self.url_input.text()

        # Run the PowerPoint script with the URL as an argument
        subprocess.run(["python3", self.genppt_path, url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setGeometry(100, 100, 600, 200)
    icon_path = os.path.join(os.getcwd(), 'icon.png')
    icon = QIcon(icon_path)
    # Set the icon for the main window
    window.setWindowIcon(icon)
    window.show()
    sys.exit(app.exec_())
