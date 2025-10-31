# coderdoc
Blood Test Parser: Data extraction from PDF or Image (via OCR) to Excel/CSV Spreadsheet with Graph Generation for Analysis

In order to get it working, set your exams files (pdf or images) inside the /exames/ folder. It can work for any language, just translate the terms inside the file compilar_exames.py file. Bellow these instructions is the reason why I made this code and attached with the files is the text that also inspired me: https://x.com/MikkaelSekeres/status/1695438539184279838/photo/1 - Thank you Mikkael Sekeres!

Explaining video: https://www.youtube.com/watch?v=bvKgxP5gXGQ

---- HOW TO RUN THE CODE DIRECTLY (WITH NO EXE COMPILATION) IN LINUX ----

Be sure you are in the coderdoc folder in the prompt, then run:

./compilar_exames.py

---- HOW TO BUILD THE EXE FOR LINUX ----

Be sure you are in the coderdoc folder in the prompt, then run:

python3 -m PyInstaller coderdoc.spec

---- HOW TO BUILD THE EXE FOR WINDOWS ----

1. Install OpenCV:

```powershell
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -m pip install opencv-python
```

If the script is currently using some extras image funcions, also install this:

```powershell
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -m pip install opencv-python-headless
```

2. Test the `cv2`:

```powershell
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -c "import cv2; print(cv2.__version__)"
```

That should print the OpenCV version (ex: `4.10.0`).


3. Just then run the EXE builder to compyle the code for Windows:

```powershell
rmdir /s /q build dist __pycache__
"C:\Users\YOUR_USERNAME\AppData\Local\Programs\Python\Python313\python.exe" -m PyInstaller coderdoc.spec
```

------------------- I BUILT THIS CODE FOR MY FATHER -------------------

After using 20 mg of Methotrexate for 3–4 years, my father developed pancytopenia.

Seven hematologists and three rheumatologists reviewed my father’s case, either through direct contact with him or via his tests. Only one rheumatologist prescribed Citoneurin (5000 mcg once daily for 3 days) and 5000 mcg once weekly for 4 weeks, which was not enough to reactivate cellular hematopoiesis as was achieved with Chronobê combined with folinic acid and active vitamin B6.

Even before considering a possible folate trap in the metabolic cascade—especially without blasts in the blood count—they suggested that my father might have myelodysplastic syndrome, without explaining what it was, its subtypes, or possible treatments, leaving our family adrift.

Few doctors listened to me as a colleague, student, programmer, English speaker, or even as a human being. This code is dedicated to them, to my family, and especially to my father.

I hope it helps others reach the correct diagnosis, just as it helped us.

Methotrexate blocked the DHFR enzyme, leading to an increase in homocysteine, which damaged the middle meninges because they have less connective tissue, are pulsatile, and are more susceptible to pressure variations. This caused an epidural hematoma that compressed areas related to the legs and silenced the CD320 and SLC19A1 genes, reducing folate and B12 entry into hematopoietic cells.

The treatment that bypassed this problem to reverse the methylation from the epigenetic alteration was folinic acid and IM chronobe. Since citoneurin has only a 1–3 day lifespan, it could not effectively enter cells via active transport (minimal in both cases due to gene silencing), nonspecific pinocytosis (5–10% increased to 10–15%), or passive diffusion (1% increased to 5%), unlike the 5–10 day window provided by hydroxocobalamin, which does not require the silenced receptors for cellular entry.

--- Portuguese version:

Após usar 20mg de Metotreaxto por 3-4 anos, meu pai desenvolveu pancitopenia. 

7 hematologistas e 3 reumatologistas revisaram o caso do meu pai diretamente em contato com ele ou exames, apenas 1 reumatologista prescreveu Citoneurin (5000 mcg 1x/dia por 3 dias) e 5000 mcg 1x/semana por 4 semanas, não sendo o suficiente para a reativação da hematopoiese celular como foi com o uso da Cronobê junto ao Ácido Folínico com Vit B6 ativa. 

Antes mesmo de verificarem um possível folate trap na cascata metabólica, especialmente sem blastos no hemograma, já sugeriram que meu pai teria síndrome mielodisplásica, não nos explicando o que seria isso, as subdiviões e os possíveis tratamentos, deixando nossa família a deriva. 

Poucos foram os médicos que me escutaram como colega, estudante, programador, falante de língua inglesa ou mesmo ser humano. Este código é dedicado a eles, à minha família e especialmente ao meu pai. 

Espero que ajude a todos a chegarem no diagnóstico verdadeiro, assim como nos ajudou.

Metotrexato bloqueou enzima dhfr, gerando aumento de homocisteína, q danificou meníngeas médias por terem menos tecido conjuntivo, serem pulsáteis e sofrerem mais variação de pressão, causando o hematoma extradural que pressionou áreas referentes às pernas e silenciamento dos genes cd320 e slc19a1, reduzindo entrada de folato e b12 na célula hematopoiética.

Tto que contornou o problema para desfazer metilação da alteração epigenética foi ácido folínico e cronobe im, visto que citoneurin possui apenas 1-3 dias de vida útil, não viabilizando entrada por transporte ativo (mínimo em ambos pois silenciados), pinocitose inespecífica (de 5-10% para 10-15%) ou difusão passiva (de 1% para 5%) como a propiciada pelos 5-10 dias da hidroxicobalamina, não demandando os receptores silenciados para a entrada celular.
