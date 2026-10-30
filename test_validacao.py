import unittest
import os
import docx
from PIL import Image

class TestValidacaoScrum(unittest.TestCase):
    def setUp(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))

    def test_arquivos_obrigatorios_existem(self):
        arquivos = [
            "Rodrigo_Ag10_REC_TI.docx",
            "infografico_scrum.png",
            "interacao_ia.png",
            "infografico.html",
            "gerar_roteiro.py"
        ]
        for arq in arquivos:
            caminho = os.path.join(self.base_dir, arq)
            self.assertTrue(os.path.exists(caminho), f"Arquivo ausente: {arq}")
            self.assertGreater(os.path.getsize(caminho), 0, f"Arquivo vazio: {arq}")

    def test_dimensoes_imagens(self):
        img_info = os.path.join(self.base_dir, "infografico_scrum.png")
        img_ia = os.path.join(self.base_dir, "interacao_ia.png")

        with Image.open(img_info) as im:
            self.assertGreater(im.width, 700)
            self.assertGreater(im.height, 500)

        with Image.open(img_ia) as im:
            self.assertGreater(im.width, 700)
            self.assertGreater(im.height, 500)

    def test_conteudo_docx(self):
        docx_path = os.path.join(self.base_dir, "Rodrigo_Ag10_REC_TI.docx")
        doc = docx.Document(docx_path)
        self.assertGreater(len(doc.paragraphs), 10)
        self.assertGreater(len(doc.tables), 0)

if __name__ == "__main__":
    unittest.main()
