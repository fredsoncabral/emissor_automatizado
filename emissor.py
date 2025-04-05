import pyautogui as pa
import time
import datetime
import os

os.chdir('C:/Users/fc/downloads/')
dir = 'C:/Users/fc/downloads/'
for f in os.listdir(dir):
    os.remove(os.path.join(dir, f))

print("============Bem-vindo===========")
print("Emissor de Relatório de Produção")
print("================================")
print()
print("Digite o nome da empresa: ")
print()

empresa = input()
x = ('http://192.11.22.01:1234/pcfui#/page/Login')
y = ('http://192.12.11.02:1234/pcfui#/page/Login')

# calculo de data
data_atual = datetime.date.today()
data_1 = data_atual - datetime.timedelta(days=1)
data_1 = [data_1.day, data_1.month, data_1.year]
dia = str(data_1[0])
mes = str(data_1[1])
ano = str(data_1[2])
dataf = dia + mes + ano
datai = ("01") + mes + ano

# tempo de pausa entre tarefas
pa.PAUSE = 1

# abrir navegador
pa.press("win")
pa.write("edge")
pa.press("enter")

# acessar pagina
if empresa == ("x"):
    pa.write(x)
else:
    pa.write(y)
pa.press("enter")
time.sleep(2)  # tempo de espera para bloco de tarefa

# logar no sistema
pa.click(x=666, y=333)  # usuario
pa.write("fr")
pa.click(x=666, y=388)  # senha
pa.click(x=666, y=488)  # enter
time.sleep(2)  # tempo de espera para bloco de tarefa

# extrair relatório
pa.click(x=1161, y=135)  # atalho relatório

if empresa == ("x"):
    pa.write("E0001")
else:
    pa.write("E0002")
pa.press("enter")
pa.click(x=383, y=172)  # parâmetros
pa.click(x=40, y=472)
pa.write(datai)  # data inicial
pa.click(x=42, y=-495)
pa.write(dataf)   # data fim
pa.click(x=40, y=586)
pa.click(x=82, y=-321)  # tipo
pa.click(x=185, y=-345)  # rodar relatório
time.sleep(15)  # tempo de espera para carregar o relatório
pa.click(x=686, y=-912)
pa.click(x=738, y=-841)  # exportar relatório excel
