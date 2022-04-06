def converter():
    def word_to_pdf(input_file, output_file):
        convert(input_file, output_file)
        print("Conversão feita com sucesso!")

    def pdf_to_word(input_file, output_file):
        arquivo_word = client.Dispatch("Word.Application")
        arquivo_word.visible=0
        arquivo_entrada = os.path.abspath(input_file)
        workbook = arquivo_word.Documents.Open(arquivo_entrada)
        arquivo_saida = os.path.abspath(output_file)
        workbook.SaveAs2(arquivo_saida, FileFormat=16)
        print("Conversão feita com sucesso!")
        if(app.questionBox("Arquivo salvo", "Deseja sair?")):
           app.stop()

    def excel_to_pdf(input_file, output_file):
        arquivo_excel = client.Dispatch("Excel.Application")
        books = arquivo_excel.Workbooks.Open(input_file).Worksheets[0]
        books.Visible = 1
        out_file = str(output_file) + ".pdf"
        books.ExportAsFixedFormat(0, out_file)
        print("Conversão feita com sucesso!")
        if(app.questionBox("Arquivo salvo", "Deseja sair?")):
           app.stop()

    def validation(input_file, dest_file, output_file):

        errors = False
        error_msgs = []
        file_formats = [".DOCX", ".XLSX", ".PDF"]
        if Path(input_file).suffix.upper() not in file_formats:
            errors = True
            error_msgs.append("Selecione um arquivo válido")

        if not(Path(dest_file)).exists():
            errors = True
            error_msgs.append("Por favor selecione um diretório válido")
        
        if len(output_file) < 1:
            errors = True
            error_msgs.append("Insira o nome do arquivo")
        
        return(errors, error_msgs)
    
    def set_file_format(input_file):
        file_format = ""
        if Path(input_file).suffix.upper() == ".XLSX":
            file_format = "xlsx"
            return file_format

        elif Path(input_file).suffix.upper() == ".DOCX":
            file_format = "docx"
            return file_format

        elif Path(input_file).suffix.upper() == ".PDF":
            file_format = "pdf"
            return file_format

    def press(button):
        if button == "Convert":
            # Modificacao necessaria
            src_file = app.getEntry("Input_File")
            dest_dir = app.getEntry("Output_Directory")
            out_file = app.getEntry("Output_name")
            errors, error_msg = validation(src_file, dest_dir, out_file)
            if errors:
                app.errorBox("Error", "\n".join(error_msg), parent=None)
            else:
                if set_file_format(src_file) == "pdf":
                    pdf_to_word(src_file, Path(dest_dir,out_file))

                elif set_file_format(src_file) == "docx":
                    word_to_pdf(src_file, Path(dest_dir,out_file))

                elif set_file_format(src_file) == "xlsx":
                    excel_to_pdf(src_file, Path(dest_dir,out_file))
        else:
            app.stop()


    app = gui("Conversor de Arquivos", useTtk=True)
    app.setTtkTheme('alt')
    app.setSize(500, 200)

    app.addLabel("Selecione o arquivo de entrada")
    app.addFileEntry("Input_File")

    app.addLabel("Select Output Directory")
    app.addDirectoryEntry("Output_Directory")

    app.addLabel("Output file name")
    app.addEntry("Output_name")

    app.addButtons(["Convert", "Quit"],press)
    app.go()


if __name__=="__main__":
    import os
    from docx2pdf import convert
    from win32com import client
    from appJar import gui
    from pathlib import Path

    converter()