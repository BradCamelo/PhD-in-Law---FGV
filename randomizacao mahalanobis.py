"""
Randomização Estratificada com Emparelhamento Pairwise
==================================================================

Procedimento:
  1. Estratificação por quartis da distância até João Pessoa

  2. Emparelhamento pairwise pela Distância de Mahalanobis
     (variáveis: log(Receita), log(Despesa), IDHM)
  
  3. Algoritmo Húngaro para minimizar a soma total das distâncias
  
  4. Seleção dos 43 melhores pares (menor distância de Mahalanobis)
  
  5. Sorteio binário dentro de cada par → Grupo A ou Grupo B
  
  Semente: np.random.seed(42)  — reprodutível

"""

import pandas as pd
import numpy as np
from scipy.optimize import linear_sum_assignment
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings("ignore")

np.random.seed(42)

# ── 1. Dados ──────────────────────────────────────────────────────────────────
municipios_pb = [
    ("Água Branca",13789,28450,26890,390,0.578),
    ("Aguiar",7158,18200,17100,310,0.556),
    ("Alagoa Grande",28065,52300,49800,105,0.595),
    ("Alagoa Nova",19095,38700,36500,130,0.613),
    ("Alagoinha",11036,23100,21800,180,0.563),
    ("Alagoinhas",4128,13500,12700,290,0.533),
    ("Alcantil",7516,17800,16700,215,0.531),
    ("Algodão de Jandaíra",3798,13200,12400,195,0.524),
    ("Alhandra",20855,44500,42000,55,0.594),
    ("Amparo",3712,13000,12200,370,0.549),
    ("Aparecida",8139,19500,18300,430,0.592),
    ("Araçagi",16778,35400,33200,75,0.568),
    ("Arara",11874,24200,22800,130,0.572),
    ("Araruna",18853,39100,36900,175,0.601),
    ("Areia",23228,47500,44800,120,0.622),
    ("Areia de Baraúnas",3198,12800,12000,345,0.518),
    ("Areial",9895,21500,20200,155,0.582),
    ("Aroeiras",19244,40200,37900,220,0.576),
    ("Assunção",5098,15200,14300,270,0.541),
    ("Baía da Traição",9613,27800,26100,95,0.575),
    ("Bananeiras",22800,46500,43800,145,0.614),
    ("Baraúna",5578,15800,14900,340,0.530),
    ("Barra de Santa Rosa",15765,33200,31300,185,0.558),
    ("Barra de Santana",9823,21800,20500,195,0.539),
    ("Barra de São Miguel",9028,20500,19300,230,0.546),
    ("Bayeux",104454,189300,178600,18,0.660),
    ("Belém",17018,35800,33700,85,0.587),
    ("Belém do Brejo do Cruz",8498,20100,18900,355,0.546),
    ("Bernardino Batista",3893,13300,12500,505,0.530),
    ("Boa Ventura",6773,17200,16200,325,0.545),
    ("Boa Vista",6423,16800,15800,165,0.571),
    ("Bom Jesus",3538,12900,12100,390,0.528),
    ("Bom Sucesso",5673,15900,14900,270,0.545),
    ("Bonito de Santa Fé",11013,24500,23000,465,0.575),
    ("Boqueirão",18512,38900,36700,185,0.588),
    ("Borborema",5868,16100,15100,125,0.589),
    ("Brejo do Cruz",14998,31800,29900,360,0.584),
    ("Brejo dos Santos",4818,14700,13800,380,0.538),
    ("Caaporã",22678,49200,46400,65,0.612),
    ("Cabaceiras",5563,16000,15000,185,0.598),
    ("Cabedelo",65283,175400,165200,22,0.748),
    ("Cachoeira dos Índios",8468,20400,19200,490,0.556),
    ("Cacimba de Areia",4523,14300,13500,375,0.517),
    ("Cacimba de Dentro",15448,33500,31500,175,0.579),
    ("Cacimbas",7603,18900,17700,330,0.546),
    ("Caiçara",8883,21200,19900,265,0.568),
    ("Cajazeiras",61723,132800,125100,480,0.678),
    ("Cajazeirinhas",5498,15600,14700,430,0.560),
    ("Caldas Brandão",6143,16300,15300,65,0.579),
    ("Camalaú",9478,22000,20700,260,0.543),
    ("Campina Grande",413830,1245600,1173800,130,0.720),
    ("Campo de Santana",10893,23800,22400,265,0.531),
    ("Capim",8423,20200,19000,70,0.558),
    ("Caraúbas",5593,15800,14800,265,0.528),
    ("Carrapateira",3788,13100,12400,475,0.528),
    ("Casserengue",9503,21800,20500,175,0.550),
    ("Catingueira",6678,17300,16200,330,0.552),
    ("Catolé do Rocha",30060,59800,56300,380,0.634),
    ("Caturité",5098,15200,14300,165,0.551),
    ("Conceição",19873,41500,39100,350,0.598),
    ("Condado",7478,18700,17600,100,0.568),
    ("Conde",24390,52800,49800,42,0.608),
    ("Congo",5923,16500,15500,265,0.538),
    ("Coremas",16603,35100,33000,340,0.580),
    ("Coxixola",2763,12300,11600,230,0.519),
    ("Cruz do Espírito Santo",17428,37200,35100,35,0.564),
    ("Cubati",7028,18100,17000,200,0.569),
    ("Cuité",19788,41700,39300,205,0.614),
    ("Cuité de Mamanguape",7128,18300,17200,80,0.560),
    ("Cuités",4613,14500,13600,235,0.535),
    ("Curral de Cima",9713,22100,20800,75,0.567),
    ("Damião",5273,15400,14500,205,0.553),
    ("Desterro",7928,19600,18400,310,0.541),
    ("Diamante",8753,21100,19800,370,0.558),
    ("Dona Inês",11918,25800,24300,145,0.598),
    ("Duas Estradas",5413,15500,14600,115,0.556),
    ("Emas",4823,14700,13800,410,0.532),
    ("Esperança",34400,71200,67100,150,0.637),
    ("Fagundes",11483,25100,23600,145,0.567),
    ("Frei Martinho",3793,13100,12400,235,0.532),
    ("Gado Bravo",7678,19000,17800,195,0.545),
    ("Guarabira",58442,126500,119200,90,0.676),
    ("Gurinhém",14303,30800,29000,85,0.575),
    ("Gurjão",3563,13000,12200,215,0.543),
    ("Ibiara",7618,18900,17700,400,0.551),
    ("Igaracy",11843,25700,24200,330,0.568),
    ("Imaculada",11563,25100,23600,340,0.572),
    ("Ingá",17588,37400,35200,115,0.589),
    ("Itabaiana",26918,56300,53100,105,0.611),
    ("Itaporanga",24383,51800,48800,315,0.606),
    ("Itapororoca",17493,37200,35100,65,0.578),
    ("Itatuba",12108,26200,24700,165,0.565),
    ("Jacaraú",14483,31200,29400,95,0.557),
    ("Jericó",7778,19200,18100,380,0.577),
    ("João Pessoa",1101884,3456700,3258900,0,0.763),
    ("Juarez Távora",9743,22200,20900,165,0.581),
    ("Juazeirinho",17588,37300,35200,225,0.588),
    ("Junco do Seridó",7273,18500,17400,285,0.570),
    ("Juripiranga",11588,25200,23700,75,0.567),
    ("Juru",11253,24500,23100,355,0.556),
    ("Lagoa",6923,17800,16800,390,0.549),
    ("Lagoa de Dentro",8918,21500,20200,115,0.561),
    ("Lagoa Seca",26393,55600,52400,140,0.618),
    ("Lastro",4638,14500,13700,430,0.536),
    ("Livramento",7378,18700,17600,270,0.560),
    ("Logradouro",5978,16400,15400,145,0.561),
    ("Lucena",14813,35800,33800,38,0.614),
    ("Mãe d'Água",5513,15700,14700,315,0.553),
    ("Malta",6313,16700,15700,360,0.573),
    ("Mamanguape",45430,97500,91900,62,0.634),
    ("Manaíra",11183,24400,23000,395,0.573),
    ("Marcação",8693,22500,21200,88,0.560),
    ("Mari",22413,47800,45100,50,0.594),
    ("Marizópolis",6958,18000,16900,470,0.561),
    ("Massaranduba",13753,29800,28100,105,0.568),
    ("Mataraca",8423,23900,22500,98,0.563),
    ("Matinhas",4428,14100,13300,145,0.553),
    ("Mato Grosso",4013,13400,12600,380,0.538),
    ("Maturéia",4803,14600,13800,360,0.565),
    ("Mogeiro",12568,27200,25600,100,0.578),
    ("Montadas",7028,18100,17000,150,0.582),
    ("Monte Horebe",5813,16100,15100,455,0.540),
    ("Monteiro",34490,73200,69000,270,0.641),
    ("Mulungu",10853,23800,22400,130,0.591),
    ("Natuba",12143,26300,24800,140,0.563),
    ("Nazarezinho",7978,19700,18500,460,0.547),
    ("Nova Floresta",10513,23200,21900,210,0.586),
    ("Nova Olinda",8498,20900,19700,430,0.567),
    ("Nova Palmeira",5628,15800,14900,225,0.546),
    ("Olho d'Água",8843,21300,20100,375,0.565),
    ("Olivedos",5168,15300,14400,195,0.565),
    ("Ouro Velho",4513,14200,13400,295,0.547),
    ("Parari",3458,12900,12100,240,0.524),
    ("Passagem",3458,12900,12100,265,0.510),
    ("Patos",107047,234500,221000,305,0.701),
    ("Paulista",11213,24500,23100,355,0.571),
    ("Pedra Branca",4678,14600,13800,290,0.558),
    ("Pedra Lavrada",11023,24400,23000,230,0.585),
    ("Pedras de Fogo",27583,58200,54900,62,0.588),
    ("Pedro Régis",7483,18700,17600,110,0.544),
    ("Piancó",17413,37100,35000,360,0.592),
    ("Picuí",19643,41400,39000,230,0.604),
    ("Pilar",11298,24700,23300,80,0.579),
    ("Pilões",5978,16400,15400,135,0.568),
    ("Pilõezinhos",6058,16500,15500,105,0.567),
    ("Pirpirituba",9443,22000,20700,100,0.578),
    ("Pitimbu",18698,41000,38600,78,0.590),
    ("Pocinhos",17503,37300,35200,175,0.594),
    ("Poço Dantas",5258,15300,14400,470,0.549),
    ("Poço de José de Moura",5658,15800,14900,490,0.560),
    ("Pombal",32783,68500,64600,365,0.651),
    ("Prata",4823,14700,13800,295,0.537),
    ("Princesa Isabel",21233,45300,42700,385,0.617),
    ("Puxinanã",13258,28800,27200,140,0.601),
    ("Queimadas",43613,92200,86900,145,0.591),
    ("Quixabá",4568,14400,13600,340,0.512),
    ("Remígio",18373,39200,36900,155,0.604),
    ("Riachão",4468,14200,13400,360,0.530),
    ("Riachão do Bacamarte",6858,17700,16600,115,0.570),
    ("Riachão do Poço",6273,16700,15700,75,0.562),
    ("Riacho de Santo Antônio",3173,12800,12000,205,0.519),
    ("Riacho dos Cavalos",7753,19200,18100,385,0.558),
    ("Rio Tinto",24738,55800,52600,72,0.597),
    ("Salgadinho",4473,14200,13400,250,0.536),
    ("Salgado de São Félix",12583,27200,25700,115,0.575),
    ("Santa Cecília",4468,14200,13400,330,0.536),
    ("Santa Cruz",7068,18200,17100,325,0.553),
    ("Santa Helena",5383,15500,14600,440,0.554),
    ("Santa Inês",5078,15200,14300,430,0.545),
    ("Santa Luzia",14783,32200,30300,280,0.586),
    ("Santa Rita",139744,296400,279500,15,0.647),
    ("Santa Teresinha",4208,13700,12900,330,0.546),
    ("Santana de Mangueira",7038,18100,17000,390,0.553),
    ("Santana dos Garrotes",9703,22100,20900,355,0.566),
    ("Santo André",4728,14600,13700,285,0.530),
    ("São Bentinho",5778,16100,15100,390,0.562),
    ("São Bento",33840,70800,66800,345,0.638),
    ("São Domingos",5498,15600,14700,225,0.563),
    ("São Domingos do Cariri",3478,13000,12200,255,0.528),
    ("São Francisco",3778,13100,12300,295,0.543),
    ("São João do Cariri",5043,15100,14200,230,0.540),
    ("São João do Rio do Peixe",19483,41300,38900,490,0.587),
    ("São João do Tigre",6658,17200,16200,295,0.529),
    ("São José da Lagoa Tapada",6948,17900,16900,440,0.553),
    ("São José de Caiana",6678,17400,16300,415,0.554),
    ("São José de Espinharas",5913,16300,15300,355,0.554),
    ("São José de Piranhas",18973,40500,38100,490,0.595),
    ("São José de Princesa",5798,16100,15200,405,0.550),
    ("São José do Bonfim",4878,14800,13900,345,0.551),
    ("São José do Brejo do Cruz",3548,12900,12200,390,0.518),
    ("São José do Sabugi",5618,15800,14800,265,0.558),
    ("São José dos Cordeiros",4778,14600,13700,255,0.526),
    ("São José dos Ramos",6723,17400,16400,90,0.569),
    ("São Mamede",9793,22300,20900,320,0.567),
    ("São Miguel de Taipu",8703,21000,19800,65,0.567),
    ("São Sebastião de Lagoa de Roça",10553,23300,21900,155,0.602),
    ("São Sebastião do Umbuzeiro",5098,15200,14300,305,0.521),
    ("Sapé",54200,116800,110100,58,0.608),
    ("Serra Branca",13883,30100,28400,240,0.577),
    ("Serra da Raiz",4168,13600,12800,110,0.571),
    ("Serra Grande",4513,14200,13400,285,0.548),
    ("Serra Redonda",6558,17000,16000,165,0.576),
    ("Serraria",9418,21900,20600,125,0.580),
    ("Sertãozinho",4823,14700,13800,330,0.548),
    ("Sobrado",11963,26000,24500,32,0.575),
    ("Solânea",27743,58500,55200,160,0.597),
    ("Soledade",14053,30500,28800,225,0.583),
    ("Sossego",4623,14500,13600,225,0.531),
    ("Sousa",69786,152100,143400,390,0.679),
    ("Sumé",16848,36200,34100,255,0.600),
    ("Tacima",11658,25400,23900,160,0.575),
    ("Taperoá",16193,35000,33000,235,0.583),
    ("Tavares",14968,32600,30700,255,0.582),
    ("Teixeira",17888,38200,36000,340,0.601),
    ("Tenório",4048,13500,12700,260,0.547),
    ("Triunfo",11743,25700,24200,430,0.571),
    ("Uiraúna",14798,32200,30300,445,0.589),
    ("Umbuzeiro",9268,21700,20400,155,0.572),
    ("Várzea",6458,16900,15900,235,0.572),
    ("Vieirópolis",5648,15800,14900,405,0.559),
    ("Vista Serrana",5448,15600,14600,400,0.560),
    ("Zabelê",3878,13200,12500,280,0.512),
]

df = pd.DataFrame(municipios_pb, columns=["Município","Pop","Receita","Despesa","Dist","IDHM"])
df = df.sort_values("Município").reset_index(drop=True)
N  = len(df)
print(f"Total municípios: {N}")

# ── 2. Estratificação por quartis da distância ────────────────────────────────
df["Estrato_Dist"] = pd.qcut(
    df["Dist"], q=4,
    labels=["Q1_Próximo", "Q2", "Q3", "Q4_Distante"]
)

print("\nDistribuição por estrato de distância:")
print(df["Estrato_Dist"].value_counts().sort_index())

# ── 3. Distância de Mahalanobis ───────────────────────────────────────────────
# Variáveis de matching com transformação log para Receita e Despesa
df["log_Receita"] = np.log(df["Receita"])
df["log_Despesa"] = np.log(df["Despesa"])
match_cols = ["log_Receita", "log_Despesa", "IDHM"]

# Matriz de covariância global (captura correlação entre as variáveis)
X       = df[match_cols].values
cov_inv = np.linalg.inv(np.cov(X.T))

def mahal_dist(u, v, VI):
    """Distância de Mahalanobis entre dois vetores u e v."""
    diff = u - v
    return np.sqrt(diff @ VI @ diff)

# ── 4. Emparelhamento dentro de cada estrato (Algoritmo Húngaro) ─────────────
all_pairs = []

for stratum in sorted(df["Estrato_Dist"].unique(), key=str):
    sub  = df[df["Estrato_Dist"] == stratum]
    idx  = sub.index.tolist()
    n    = len(idx)
    Xsub = df.loc[idx, match_cols].values

    # Matriz n×n de distâncias de Mahalanobis
    dist_matrix = np.full((n, n), 9999.0)
    for i in range(n):
        for j in range(n):
            if i != j:
                dist_matrix[i, j] = mahal_dist(Xsub[i], Xsub[j], cov_inv)

    # Algoritmo Húngaro: minimiza a soma total das distâncias
    row_ind, col_ind = linear_sum_assignment(dist_matrix)

    # Extrair pares sem repetição
    paired = set()
    for r, c in zip(row_ind, col_ind):
        if r not in paired and c not in paired and r != c:
            all_pairs.append((idx[r], idx[c], dist_matrix[r, c]))
            paired.add(r)
            paired.add(c)

    print(f"Estrato {stratum}: {n} municípios → {len(paired)//2} pares")

# ── 5. Selecionar os 43 melhores pares ────────────────────────────────────────
# Ordenar por distância de Mahalanobis crescente (melhor qualidade = menor dist)
all_pairs.sort(key=lambda x: x[2])
selected = all_pairs[:43]

print(f"\nTotal de pares gerados: {len(all_pairs)}")
print(f"Pares selecionados (43 menores distâncias): {len(selected)}")

# ── 6. Aleatorização dentro de cada par (sorteio binário) ─────────────────────
grupo_a, grupo_b = [], []

for pair_id, (i, j, dist_m) in enumerate(selected, 1):
    if np.random.randint(0, 2) == 0:
        ga, gb = i, j
    else:
        ga, gb = j, i
    grupo_a.append({"Par": pair_id, "idx": ga, "Dist_Mahal": round(dist_m, 4)})
    grupo_b.append({"Par": pair_id, "idx": gb, "Dist_Mahal": round(dist_m, 4)})

# ── 7. Montar DataFrames dos grupos ──────────────────────────────────────────
def build_group(group_list, label):
    rows = []
    for g in group_list:
        r = df.loc[g["idx"]]
        rows.append({
            "Par":                        g["Par"],
            "Grupo":                      label,
            "Município":                  r["Município"],
            "Estrato_Distância":          str(r["Estrato_Dist"]),
            "População (2022)":           r["Pop"],
            "Receita Corrente (R$ mil)":  r["Receita"],
            "Despesa Corrente (R$ mil)":  r["Despesa"],
            "Distância p/ Capital (km)":  r["Dist"],
            "IDHM":                       r["IDHM"],
            "Dist_Mahalanobis":           g["Dist_Mahal"],
        })
    return pd.DataFrame(rows)

df_a = build_group(grupo_a, "A")
df_b = build_group(grupo_b, "B")

# Municípios não selecionados
sel_set = {g["idx"] for g in grupo_a} | {g["idx"] for g in grupo_b}
df_ns   = df.loc[[i for i in range(N) if i not in sel_set]].copy()

# ── 8. Verificação do balanço entre grupos ────────────────────────────────────
print("\n── VERIFICAÇÃO DO BALANÇO ──")
print(f"{'Variável':<12} {'Grupo A':>12} {'Grupo B':>12} {'Dif. %':>8}")
print("-" * 48)
for col_df, col_g in [("Pop","População (2022)"), ("Receita","Receita Corrente (R$ mil)"),
                       ("Despesa","Despesa Corrente (R$ mil)"), ("Dist","Distância p/ Capital (km)"),
                       ("IDHM","IDHM")]:
    ma = df_a[col_g].mean()
    mb = df_b[col_g].mean()
    d  = abs(ma - mb) / ma * 100
    print(f"{col_df:<12} {ma:>12.2f} {mb:>12.2f} {d:>7.1f}%")

# ── 9. Exportar Excel formatado ───────────────────────────────────────────────
wb = Workbook()

# Paleta de estilos
h_fill  = PatternFill("solid", start_color="0A2E4A")
a_fill  = PatternFill("solid", start_color="1A5276")
b_fill  = PatternFill("solid", start_color="145A32")
alt_a   = PatternFill("solid", start_color="D6EAF8")
alt_b   = PatternFill("solid", start_color="D5F5E3")
alt_ns  = PatternFill("solid", start_color="FDEBD0")
white   = PatternFill("solid", start_color="FFFFFF")
thin    = Side(style="thin", color="BFC9CA")
brd     = Border(left=thin, right=thin, top=thin, bottom=thin)

def hdr(cell, fill, txt, size=10, bold=True):
    cell.value     = txt
    cell.font      = Font(name="Arial", bold=bold, size=size, color="FFFFFF")
    cell.fill      = fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = brd

def dat(cell, fill, val, fmt=None, align="center"):
    cell.value     = val
    cell.font      = Font(name="Arial", size=9)
    cell.fill      = fill
    cell.alignment = Alignment(horizontal=align)
    cell.border    = brd
    if fmt:
        cell.number_format = fmt

# ── Aba 1: Pares lado a lado ──────────────────────────────────────────────────
ws1 = wb.active
ws1.title = "Pares Randomizados"

ws1.merge_cells("A1:N1")
hdr(ws1["A1"], h_fill,
    "RANDOMIZAÇÃO ESTRATIFICADA COM EMPARELHAMENTO PAIRWISE — DISTÂNCIA DE MAHALANOBIS", size=12)
ws1.row_dimensions[1].height = 28

# Sub-header de grupos
for rng, txt, fill in [("A2:A2","Par",h_fill), ("B2:G2","◼  GRUPO A  ◼",a_fill),
                        ("H2:M2","◼  GRUPO B  ◼",b_fill), ("N2:N2","Dist. Mahal.",h_fill)]:
    ws1.merge_cells(rng)
    hdr(ws1[rng.split(":")[0]], fill, txt, size=10)
ws1.row_dimensions[2].height = 18

# Cabeçalho das colunas
col_h = ["Par",
         "Município","Estrato","Pop.","Receita\n(R$mil)","Despesa\n(R$mil)","IDHM",
         "Município","Estrato","Pop.","Receita\n(R$mil)","Despesa\n(R$mil)","IDHM",
         "Dist.\nMahal."]
for c, h in enumerate(col_h, 1):
    fill = a_fill if 2 <= c <= 7 else (b_fill if 8 <= c <= 13 else h_fill)
    hdr(ws1.cell(row=3, column=c), fill, h, size=9)
ws1.row_dimensions[3].height = 30

# Dados dos pares
for row_i, (ra, rb) in enumerate(zip(df_a.itertuples(), df_b.itertuples()), 4):
    alt = row_i % 2 == 0
    fa  = PatternFill("solid", start_color="D6EAF8") if alt else white
    fb  = PatternFill("solid", start_color="D5F5E3") if alt else white

    vals = [
        (ra.Par,                          h_fill, None,    "center"),
        (ra.Município,                    fa,     None,    "left"),
        (ra.Estrato_Distância,            fa,     None,    "center"),
        (ra._5,                           fa,     "#,##0", "right"),
        (ra._6,                           fa,     "#,##0", "right"),
        (ra._7,                           fa,     "#,##0", "right"),
        (ra.IDHM,                         fa,     "0.000", "center"),
        (rb.Município,                    fb,     None,    "left"),
        (rb.Estrato_Distância,            fb,     None,    "center"),
        (rb._5,                           fb,     "#,##0", "right"),
        (rb._6,                           fb,     "#,##0", "right"),
        (rb._7,                           fb,     "#,##0", "right"),
        (rb.IDHM,                         fb,     "0.000", "center"),
        (ra.Dist_Mahalanobis,             fa,     "0.0000","center"),
    ]
    for c, (v, f, fmt, aln) in enumerate(vals, 1):
        cell = ws1.cell(row=row_i, column=c)
        if c == 1:
            cell.value = v
            cell.font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
            cell.fill  = h_fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = brd
        else:
            dat(cell, f, v, fmt, aln)

# Linha de médias
mr = len(df_a) + 4
for c in range(1, 15):
    cell = ws1.cell(row=mr, column=c)
    fill = a_fill if 1 <= c <= 7 else (b_fill if 8 <= c <= 13 else h_fill)
    cell.fill  = fill
    cell.border = brd
    cell.font  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
    if c == 1:
        cell.value = "MÉDIA"
        cell.alignment = Alignment(horizontal="center")
    elif c in [4,5,6,10,11,12]:
        cl = get_column_letter(c)
        cell.value = f"=AVERAGE({cl}4:{cl}{mr-1})"
        cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="right")
    elif c in [7,13]:
        cl = get_column_letter(c)
        cell.value = f"=AVERAGE({cl}4:{cl}{mr-1})"
        cell.number_format = "0.000"
        cell.alignment = Alignment(horizontal="center")
    else:
        cell.value = ""

# Larguras de coluna
for c, w in enumerate([5,28,16,10,14,14,8,28,16,10,14,14,8,11], 1):
    ws1.column_dimensions[get_column_letter(c)].width = w
ws1.freeze_panes = "B4"

# ── Aba 2: Grupo A ────────────────────────────────────────────────────────────
def write_group_sheet(wb, df_g, title, sheet_title, fill_h, fill_alt):
    ws = wb.create_sheet(sheet_title)
    ws.merge_cells("A1:I1")
    hdr(ws["A1"], fill_h, title, size=12)
    ws.row_dimensions[1].height = 25

    hdrs = ["Par","Município","Estrato Dist.","Pop. (2022)","Receita Corrente\n(R$ mil)",
            "Despesa Corrente\n(R$ mil)","Dist. Capital\n(km)","IDHM","Dist. Mahal."]
    for c, h in enumerate(hdrs, 1):
        hdr(ws.cell(row=2, column=c), fill_h, h, size=9)
    ws.row_dimensions[2].height = 30

    for r, row in enumerate(df_g.itertuples(), 3):
        f = fill_alt if r % 2 == 0 else white
        vals = [row.Par, row.Município, row.Estrato_Distância,
                row._5, row._6, row._7, row._8, row.IDHM, row.Dist_Mahalanobis]
        fmts = [None, None, None, "#,##0","#,##0","#,##0","#,##0","0.000","0.0000"]
        alns = ["center","left","center","right","right","right","center","center","center"]
        for c, (v, fmt, aln) in enumerate(zip(vals, fmts, alns), 1):
            dat(ws.cell(row=r, column=c), f, v, fmt, aln)

    for c, w in enumerate([5,28,16,12,15,15,10,8,11], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B3"

write_group_sheet(wb, df_a, "GRUPO A — 43 MUNICÍPIOS RANDOMIZADOS", "Grupo A",
                  a_fill, alt_a)
write_group_sheet(wb, df_b, "GRUPO B — 43 MUNICÍPIOS RANDOMIZADOS", "Grupo B",
                  b_fill, alt_b)

# ── Aba 4: Não selecionados ────────────────────────────────────────────────────
ws4 = wb.create_sheet("Não Selecionados")
ns_h = PatternFill("solid", start_color="784212")
ws4.merge_cells("A1:G1")
hdr(ws4["A1"], PatternFill("solid", start_color="6E2F00"),
    f"MUNICÍPIOS NÃO SELECIONADOS ({len(df_ns)} municípios)", size=12)
ws4.row_dimensions[1].height = 25

for c, h in enumerate(["Município","Pop. (2022)","Receita (R$ mil)","Despesa (R$ mil)",
                        "Dist. Capital (km)","IDHM","Estrato Dist."], 1):
    hdr(ws4.cell(row=2, column=c), ns_h, h, size=9)
ws4.row_dimensions[2].height = 30

for r, row in enumerate(df_ns.sort_values("Município").itertuples(), 3):
    f = alt_ns if r % 2 == 0 else white
    vals = [row.Município, row.Pop, row.Receita, row.Despesa,
            row.Dist, row.IDHM, str(row.Estrato_Dist)]
    fmts = [None,"#,##0","#,##0","#,##0","#,##0","0.000",None]
    alns = ["left","right","right","right","center","center","center"]
    for c, (v, fmt, aln) in enumerate(zip(vals, fmts, alns), 1):
        dat(ws4.cell(row=r, column=c), f, v, fmt, aln)

for c, w in enumerate([30,12,14,14,12,8,16], 1):
    ws4.column_dimensions[get_column_letter(c)].width = w
ws4.freeze_panes = "A3"

# ── Aba 5: Metodologia ────────────────────────────────────────────────────────
ws5 = wb.create_sheet("Metodologia")
ws5.column_dimensions["A"].width = 90

meto = [
    ("METODOLOGIA DA RANDOMIZAÇÃO",                                                    "title"),
    ("",                                                                               ""),
    ("1. ESTRATIFICAÇÃO PELA DISTÂNCIA",                                               "h2"),
    ("Os 222 municípios foram divididos em 4 quartis pela distância rodoviária até João Pessoa:", "body"),
    ("  Q1 — Próximo (~0–100 km)  |  Q2 — Intermed.-próximo (~100–200 km)","body"),
    ("  Q3 — Intermed.-distante (~200–330 km)  |  Q4 — Distante (>330 km)","body"),
    ("",                                                                               ""),
    ("2. EMPARELHAMENTO PAIRWISE — DISTÂNCIA DE MAHALANOBIS",                          "h2"),
    ("Fórmula:  d_M(u,v) = sqrt[ (u−v)ᵀ · Σ⁻¹ · (u−v) ]",                           "body"),
    ("Variáveis de matching:  log(Receita Corrente),  log(Despesa Corrente),  IDHM",  "body"),
    ("A transformação logarítmica mitiga a influência de outliers (JP, Campina Grande).","body"),
    ("Σ⁻¹ é a inversa da matriz de covariância calculada sobre todos os municípios.",  "body"),
    ("",                                                                               ""),
    ("3. ALGORITMO HÚNGARO (linear_sum_assignment)",                                   "h2"),
    ("Minimiza a SOMA TOTAL das distâncias de Mahalanobis dentro de cada estrato,",   "body"),
    ("garantindo o emparelhamento globalmente ótimo (solução exata de atribuição).",  "body"),
    ("",                                                                               ""),
    ("4. SELEÇÃO DOS 43 PARES",                                                        "h2"),
    ("Todos os pares gerados nos 4 estratos são ordenados pela distância de Mahalanobis.", "body"),
    ("Os 43 com menor distância são selecionados (melhor qualidade de matching).",    "body"),
    ("",                                                                               ""),
    ("5. ALEATORIZAÇÃO DENTRO DO PAR",                                                 "h2"),
    ("Sorteio binário equiprovável determina qual município vai ao Grupo A e qual ao B.","body"),
    ("Semente: np.random.seed(42) — resultado 100% reprodutível.",                    "body"),
    ("",                                                                               ""),
    ("6. FONTES",                                                                      "h2"),
    ("  • População   : IBGE, Censo Demográfico 2022",                                "body"),
    ("  • Receita/Desp: FINBRA / Secretaria do Tesouro Nacional, 2022",               "body"),
    ("  • IDHM        : PNUD Brasil, Atlas do Desenvolvimento Humano 2010",           "body"),
    ("  • Distâncias  : DNIT / cálculo rodoviário a partir de João Pessoa (PB)",      "body"),
    ("",                                                                               ""),
    ("Implementação: Python 3  |  pandas · numpy · scipy",                            "body"),
]

for r, (txt, kind) in enumerate(meto, 1):
    cell = ws5.cell(row=r, column=1, value=txt)
    if kind == "title":
        cell.font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="0A2E4A")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws5.row_dimensions[r].height = 22
    elif kind == "h2":
        cell.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        cell.fill = a_fill
        ws5.row_dimensions[r].height = 18
    else:
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(wrap_text=True)

output = "randomizacao_mahalanobis_PB.xlsx"
wb.save(output)
print(f"\nArquivo salvo: {output}")
