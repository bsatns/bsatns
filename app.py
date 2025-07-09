from flask import Flask, render_template, request, send_file
import pandas as pd
import os 
from werkzeug.utils import secure_filename
import zipfile 

app = Flask (__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route ('/', methods = ['GET', 'POST'])
def index():
    if request.method =='POST':
        num_parts= int (request.form ['num_parts'])
        formatar = request.form.get ('formatar') == 'sim'

        arquivos = request.files.getlist ('planilhas')
        print ('Arquivo(s) recebido(s): ', arquivos)

        if not arquivos:
            return 'nenhum arquivo anexado', 400
        
        os.makedirs (RESULT_FOLDER, exist_ok=True)
        zip_path = os.path.join (RESULT_FOLDER, 'Planilhas_divididas.zip')

        with zipfile.ZipFile (zip_path, 'w') as zipf:
            for file in arquivos: 
                filename= secure_filename (file.filename)
                path = os.path.join (UPLOAD_FOLDER, filename)
                file.save (path)

                df = pd.read_excel (path)
                tamanho = len(df) // num_parts + (len(df) % num_parts > 0)

                for i in range (num_parts):
                    part_df = df[i*tamanho: (i+1)* tamanho]
                    part_filename = f"{filename.rsplit('.', 1)[0]}_part{i+10}.xlsx"
                    part_path = os.path.join(UPLOAD_FOLDER, part_filename)
                    part_df.to_excel(part_path, index=False)
                    zipf.write(part_path, arcname=part_filename)
                
            
            return send_file(zip_path, as_attachment=True)
        
        if not arquivos:
            return "Nenhum arquivo anexado", 400  
    
    return render_template ("index.html")
if __name__ == '__main__':
    app.run(debug=True)
