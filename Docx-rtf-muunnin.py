import sys
import subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QAction
from PyQt5.QtCore import QTimer
from docx import Document
from striprtf.striprtf import rtf_to_text


class WordViewer(QMainWindow):
    def __init__(self, file_to_open=None):
        super().__init__()

        # Pääikkunan asetukset
        self.setWindowTitle("Word/RTF Viewer")
        self.setGeometry(100, 100, 800, 600)

        # Luo valikko
        self.create_menu()

        # Jos sovellus käynnistetään tiedostopolulla, käsittele se
        if file_to_open:
            self.handle_file(file_to_open)
            # Ajasta sovelluksen sulkeminen
            QTimer.singleShot(1000, self.cleanup_and_exit)

    def cleanup_and_exit(self):
        """Siivoa ja sulje sovellus"""
        QApplication.quit()

    def create_menu(self):
        # Valikko tiedoston avaamiseen
        menu = self.menuBar()
        file_menu = menu.addMenu("Tiedosto")

        # Luo "Avaa" -toiminto
        open_action = QAction("Avaa Word- tai RTF-dokumentti", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        # Luo "Poistu" -toiminto
        exit_action = QAction("Poistu", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def open_file(self):
        # Avaa tiedostovalitsin
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Avaa Word- tai RTF-dokumentti", "", "Supported Files (*.docx *.rtf)"
        )
        if file_path:
            self.handle_file(file_path)

    def handle_file(self, file_path):
        # Tarkista tiedoston tyyppi ja käsittele sen mukaisesti
        _, file_extension = os.path.splitext(file_path)
        if file_extension.lower() == ".docx":
            self.convert_word_to_txt(file_path)
        elif file_extension.lower() == ".rtf":
            self.convert_rtf_to_txt(file_path)
        else:
            print("Tiedostotyyppiä ei tueta.")

    def convert_word_to_txt(self, file_path):
        try:
            # Lue Word-dokumentti
            document = Document(file_path)
            text = ""

            # Muunna sisältö tavalliseksi tekstiksi
            for paragraph in document.paragraphs:
                text += f"{paragraph.text}\n"

            # Tallenna teksti samaan kansioon txt-muodossa
            self.save_text_file(file_path, text)
        except Exception as e:
            print(f"Virhe Word-tiedostoa käsiteltäessä: {e}")

    def convert_rtf_to_txt(self, file_path):
        try:
            # Lue RTF-tiedosto ja muunna tavalliseksi tekstiksi
            with open(file_path, "r", encoding="utf-8") as rtf_file:
                rtf_content = rtf_file.read()
                text = rtf_to_text(rtf_content)

            # Tallenna teksti samaan kansioon txt-muodossa
            self.save_text_file(file_path, text)
        except Exception as e:
            print(f"Virhe RTF-tiedostoa käsiteltäessä: {e}")

    def save_text_file(self, original_file_path, text):
        try:
            # Määritä uuden tiedoston polku ja nimi
            base_name = os.path.splitext(original_file_path)[0]  # Ilman tiedostopäätettä
            txt_file_path = f"{base_name}.txt"

            # Tallenna teksti tiedostoon
            with open(txt_file_path, "w", encoding="utf-8") as f:
                f.write(text)

            print(f"Tekstitiedosto tallennettu: {txt_file_path}")

            # Avaa tallennettu tiedosto Muistiossa
            subprocess.Popen(["notepad.exe", txt_file_path])
        except Exception as e:
            print(f"Virhe tekstitiedostoa tallennettaessa: {e}")


def main():
    # Tarkista, onko sovellus käynnistetty tiedostopolulla
    file_to_open = sys.argv[1] if len(sys.argv) > 1 else None

    app = QApplication(sys.argv)
    viewer = WordViewer(file_to_open)
    
    # Näytä ikkuna vain jos ei ole tiedostoa käsiteltävänä
    if not file_to_open:
        viewer.show()
        sys.exit(app.exec_())
    else:
        # Jos tiedosto käsitellään, käynnistä sovelluksen pääsilmukka
        # mutta älä näytä ikkunaa
        app.exec_()


if __name__ == "__main__":
    main()
