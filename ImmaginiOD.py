import ftplib
import os
import socket
import sys

prefis_image = 'CANC'

def download_file(_ftp, _nome_file):
    try:
        with open("C:/JOIA/t-shop/FOTO/ARTICOLI_OFFICE_ADOK/new/"+prefis_image+_nome_file, "wb") as file:
            _ftp.retrbinary("RETR " + _nome_file, file.write)
            print("File scaricato:", _nome_file)
    except ConnectionResetError:
        print("Errore di connessione durante il download di", _nome_file, "- Riprovo")
        _ftp.quit()
        _ftp = ftplib.FTP("ftp.#######.eu")
        _ftp.login(user="#######", passwd="#######")
        _ftp.cwd("/MD")
        download_file(_ftp, _nome_file)

ftp = ftplib.FTP("ftp.#######.eu")

try:
    ftp.login(user="#######", passwd="#######")
    print("CONNESSIONE FTP OK")
except ftplib.error_perm as e:
    print("AUTENTICAZIONE SERVER FTP FALLITA:", e)
    os.system("pause")
    sys.exit()
except ftplib.error_proto as e:
    print("PROTOCOLLO DI CONNESSIONE ERRATO:", e)
    os.system("pause")
    sys.exit()
except socket.error as e:
    print("ERRORE DI CONNESSIONE:", e)
    os.system("pause")
    sys.exit()

try:
    ftp.cwd("/MD")
    print("CWD /MD OK")
except ftplib.error_perm as e:
    print("CAMBIO DIRECTORY IN /MD FALLITO")
    os.system("pause")
    sys.exit()
except ftplib.error_proto as e:
    print("PROTOCOLLO DI CONNESSIONE ERRATO CWD:", e)
    os.system("pause")
    sys.exit()
except socket.error as e:
    print("ERRORE DI CONNESSIONE CWD:", e)
    os.system("pause")
    sys.exit()

try:
    elenco_files = ftp.nlst()
except ftplib.error_perm as e:
    print("ERRORE COMANDO NLST:", e)
    os.system("pause")
    sys.exit()
except ftplib.error_proto as e:
    print("PROTOCOLLO DI CONNESSIONE ERRATO NLST", e)
    os.system("pause")
    sys.exit()
except socket.error as e:
    print("ERRORE DI CONNESSIONE NLST", e)
    os.system("pause")
    sys.exit()

total_files = len(elenco_files)
files_downloaded = 0

for nome_file in elenco_files:
    download_file(ftp, nome_file)

ftp.quit()
print("Download completato")
os.system("pause")
sys.exit()
