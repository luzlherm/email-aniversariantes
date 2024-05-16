import win32com.client as win32

outlook = win32.Dispatch("outlook.application")

emailOutlook = outlook.CreateItem(0)

emailOutlook.To = "ana@gmail.com"
emailOutlook.Subject = "Feliz Aniversário"
emailOutlook.HTMLBody = """
<p>Parabéns, Ana!</p>
<p>Esse é um dia especial, aproveite seu dia!</p>
<p>Atenciosamente.</p>
"""

#save - Salvar como rascunho / draft
#send - envia
emailOutlook.save()