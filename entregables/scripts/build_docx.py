from docx_lib import Doc
from contenido1 import parte1
from contenido2 import parte2
D = Doc(); parte1(D); parte2(D)
D.save('../Contenido_Nuevo_Tesis_Cosecha_Viva.docx'); print('tablas', D.nt, 'figuras', D.nf)
