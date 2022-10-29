Baixar Anexos do Outlook
=====

É um aplicativo feito no python utilizando PyQt5, para a interface gráfica, e win32com.client, para conectar no Outlook do Windows.
Ele foi feito com o objetivo de baixar todos os anexos de acordo com os filtros nos e-mails do Outlook.

![app](https://github.com/itff/baixar_anexo_outlook/blob/main/images/mainwindow.PNG?raw=true)

Como Executar
===================

Tem duas formas:
- Abra o executável App_Baixar_Anexos_Email.exe
- Baixe a pasta code e execute no Python o app.py

Como criar o Executável
===================

- Abra o Anaconda Prompt
- Mude o diretório para a pasta com o app.py com o *cd C:\\*
- Rode:
```
pyinstaller --name="App_Baixar_Anexos_Email" --windowed --onefile app.py
```

*Precisa ter o Anaconda instalado.*
