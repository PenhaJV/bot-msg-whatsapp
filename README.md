# Envio Automático de Mensagens via WhatsApp Web

## Descrição

Este script automatiza o processo de envio de mensagens via WhatsApp Web utilizando uma lista de contatos armazenada em uma planilha Excel. O script utiliza as bibliotecas `openpyxl` para manipulação da planilha, `pyautogui` para automação de cliques na tela e `webbrowser` para abrir os links do WhatsApp Web.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `openpyxl`
  - `pyautogui`

Você pode instalar as bibliotecas necessárias com o seguinte comando:

```bash
pip install openpyxl pyautogui
```

## Como Usar

1. Certifique-se de que o WhatsApp Web esteja configurado e que você possa fazer login manualmente.

2. A pasta `data` já contém a imagem `seta_whatsapp.png` e a planilha `contatos.xlsx`. Utilize essa planilha e preencha os campos **Nome** (Coluna A) e **Telefone** (Coluna B). 

3. Execute o script `main.py`:

```bash
python main.py
```

4. O script abrirá o WhatsApp Web, enviará as mensagens e registrará quaisquer erros no arquivo `erros.csv`.

## Considerações

- O tempo de espera (`sleep`) pode precisar ser ajustado dependendo da velocidade da sua conexão de internet e do desempenho do seu computador.
- O script tenta enviar a mensagem e fechar a aba do navegador automaticamente após cada envio.

## Licença

Este projeto está licenciado sob os termos da licença MIT.