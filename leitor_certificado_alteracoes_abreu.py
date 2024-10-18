import ssl
import socket
import pandas as pd
from OpenSSL import crypto

def get_cert_info(domain):
    try:
        ############  # Estabelecer conexão para obter o certificado ############
        ctx = ssl.create_default_context()
        with ctx.wrap_socket(socket.socket(), server_hostname=domain) as s:
            s.connect((domain, 443))
            cert = s.getpeercert(True)
            x509 = crypto.load_certificate(crypto.FILETYPE_ASN1, cert)

        ######### # Obter o Common Name (CN) ############
        subject = x509.get_subject()
        cn = subject.CN

        ############ # Obter o Subject Alternative Names (SAN) # ###########
        alt_names = []
        for i in range(x509.get_extension_count()):
            ext = x509.get_extension(i)
            if ext.get_short_name() == b'subjectAltName':
                alt_names = str(ext).split(", ")

        # ########### # Obter a data de expiração # ########### #
        expiry_date = x509.get_notAfter().decode('utf-8')

        return {'CN': cn, 'SAN': alt_names, 'Expiry Date': expiry_date}

    except Exception as e:
        print(f"Erro ao obter informações do certificado para {domain}: {e}")
        return {'CN': 'Error', 'SAN': 'Error', 'Expiry Date': 'Error', 'Error': str(e)}

def export_to_excel(data, filename='certificates.xlsx'):
    df = pd.DataFrame(data)
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Certificates', index=False)

# ########### Lista de sites # ###########
sites = ['itcsecurity.com.br', 'invalidsite.com']  # Exemplo com site inválido para testar erros

# ########### Coletar dados dos certificados # ###########
cert_data = []
for site in sites:
    info = get_cert_info(site)
    cert_data.append({
        'Common Name (CN)': info['CN'],
        'Subject Alt Names (SAN)': ", ".join(info['SAN']),
        'Expiry Date': info['Expiry Date'],
        'Error': info.get('Error', '')
    })

# Verificando se houve erros na coleta de dados
if any(item['Error'] for item in cert_data):
    print("Ocorreram erros ao coletar dados de alguns certificados.")
    for item in cert_data:
        if item['Error']:
            print(f"Erro para {item['Common Name (CN)']}: {item['Error']}")

# ########### Exportar para Excel # ###########
try:
    export_to_excel(cert_data, 'certificates.xlsx')
    print("Arquivo Excel criado com sucesso!")
except Exception as e:
    print(f"Erro ao exportar para Excel: {e}")
