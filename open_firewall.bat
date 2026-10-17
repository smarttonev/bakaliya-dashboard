@echo off
chcp 65001 >nul
echo Разрешаване на порт 8501 във Windows Firewall за достъп от други компютри...
netsh advfirewall firewall delete rule name="Streamlit Mistral 8501" >nul 2>&1
netsh advfirewall firewall add rule name="Streamlit Mistral 8501" dir=in action=allow protocol=TCP localport=8501 profile=any
if errorlevel 1 (
    echo [ГРЕШКА] Моля, стартирайте този файл като Администратор (Десен бутон -> Run as administrator)!
) else (
    echo [УСПЕХ] Порт 8501 е разрешен успешно във Windows Firewall!
    echo Другите компютри вече могат да отварят: http://192.168.5.99:8501
)
pause
