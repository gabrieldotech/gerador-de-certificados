from docx import Document
from docx2pdf import convert
import os
import webbrowser
from datetime import datetime

def gerar_certificado(nome_aluno, nome_professor):
    current_dir = os.getcwd()
    print(f"Diretório atual: {current_dir}")

    caminho_template = 'template_certificado.docx'
    
    try:
        doc = Document(caminho_template)

        for p in doc.paragraphs:
            if 'NOME_DO_ALUNO' in p.text:
                p.text = p.text.replace('NOME_DO_ALUNO', nome_aluno)
            if 'DATA_DE_EMISSAO' in p.text:
                data_emissao = datetime.now().strftime('%d/%m/%Y')
                p.text = p.text.replace('DATA_DE_EMISSAO', data_emissao)
            if 'NOME_DO_PROFESSOR' in p.text:
                p.text = p.text.replace('NOME_DO_PROFESSOR', nome_professor)

        pasta_saida = os.path.join(current_dir, 'certificados')
        if not os.path.exists(pasta_saida):
            os.makedirs(pasta_saida) 
        
        output_word = os.path.join(pasta_saida, f'certificado_{nome_aluno}.docx')
        doc.save(output_word)
        
        print(f'Convertendo o documento para PDF...')
        convert(output_word)
        
        output_pdf = output_word.replace(".docx", ".pdf")
        print(f'Certificado gerado e salvo como {output_pdf}')
    
        webbrowser.open_new(f'file:///{os.path.abspath(output_pdf)}')
    except Exception as e:
        print(f"Ocorreu um erro ao gerar o certificado: {e}")

gerar_certificado('João da Silva', 'Professor Gabriel')
