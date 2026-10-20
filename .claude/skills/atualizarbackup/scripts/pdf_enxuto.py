# PDF mínimo, só texto ASCII (acentos como escapes octais WinAnsi), para subir pelo conector do Drive como texto.
import sys
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.utils import simpleSplit
sys.path.insert(0, sys.argv[1])
from conteudo import SECOES, DATA

W, H, ML, MT, MB = 595, 842, 50, 60, 50
LARG = W - 2 * ML

def esc(s):
    out = []
    for ch in s.encode("cp1252"):
        c = chr(ch)
        if c in "()\\": out.append("\\" + c)
        elif ch > 126: out.append("\\%03o" % ch)
        else: out.append(c)
    return "".join(out)

linhas = []  # (fonte, tamanho, cor, x, texto, espaco_antes)
def add(txt, f="F1", sz=9, cor="0 g", x=0, antes=0, larg=LARG):
    partes = simpleSplit(txt, "Helvetica-Bold" if f == "F2" else "Helvetica", sz, larg - x)
    for i, p in enumerate(partes):
        linhas.append((f, sz, cor, x, p, antes if i == 0 else 0))

AZ = "0.12 0.23 0.37 rg"
add("Contexto para continuar o trabalho", "F2", 18, AZ)
add("Matheus Macedo - Coordenador de Frota e Administrativo, LGR (Locadora Grillo e Ribeiro). Retrato de " + DATA + ".", "F2", 10, AZ, antes=4)
for tit, blocos in SECOES:
    add(tit, "F2", 12.5, AZ, antes=12)
    for tipo, val in blocos:
        if tipo == "p": add(val, antes=3)
        elif tipo == "h": add(val, "F2", 10, AZ, antes=7)
        elif tipo == "ul":
            for it in val:
                partes = simpleSplit(it, "Helvetica", 9, LARG - 12)
                for i, p in enumerate(partes):
                    linhas.append(("F1", 9, "0 g", 12, ("•  " if i == 0 else "") + p, 2.5 if i == 0 else 0))
                    if i == 0: linhas[-1] = ("F1", 9, "0 g", 3, "•  " + p, 2.5)
        elif tipo == "t":
            cab = val[0]
            for r in val[1:]:
                if cab[0] == "#":
                    add(r[0] + ". " + r[1] + "  [" + r[2] + "]", antes=3)
                else:
                    add(r[0], "F2", 9, "0 g", antes=4)
                    for c, v in zip(cab[1:], r[1:]):
                        add((c + ": " if len(cab) > 2 else "") + v, x=12)

# paginação
paginas, atual, y = [], [], H - MT
for f, sz, cor, x, t, antes in linhas:
    lead = sz * 1.32
    if y - antes - lead < MB:
        paginas.append(atual); atual, y = [], H - MT; antes = 0
    y -= antes + lead
    atual.append((f, sz, cor, x, t, y))
paginas.append(atual)

objs = []
objs.append("<< /Type /Catalog /Pages 2 0 R >>")
n = len(paginas)
kids = " ".join("%d 0 R" % (5 + 2 * i) for i in range(n))
objs.append("<< /Type /Pages /Count %d /Kids [ %s ] >>" % (n, kids))
objs.append("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
objs.append("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
for i, pg in enumerate(paginas):
    s = ["BT"]
    estado = None
    for f, sz, cor, x, t, y in pg:
        if (f, sz, cor) != estado:
            s.append("/%s %g Tf %s" % (f, sz, cor)); estado = (f, sz, cor)
        s.append("1 0 0 1 %d %.1f Tm (%s) Tj" % (ML + x, y, esc(t)))
    s.append("/F1 7 Tf 0.5 g 1 0 0 1 %d 30 Tm (%s) Tj" % (ML, esc("Contexto Matheus / LGR - %s - pág. %d de %d" % (DATA, i + 1, n))))
    s.append("ET")
    stream = "\n".join(s)
    objs.append("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents %d 0 R >>" % (W, H, 6 + 2 * i))
    objs.append("<< /Length %d >>\nstream\n%s\nendstream" % (len(stream), stream))

out = "%PDF-1.4\n"
offs = []
for k, o in enumerate(objs):
    offs.append(len(out))
    out += "%d 0 obj\n%s\nendobj\n" % (k + 1, o)
xref = len(out)
out += "xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
out += "".join("%010d 00000 n \n" % o for o in offs)
out += "trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
assert all(ord(c) < 128 for c in out)
open(sys.argv[2], "w").write(out)
print(len(out), n)
