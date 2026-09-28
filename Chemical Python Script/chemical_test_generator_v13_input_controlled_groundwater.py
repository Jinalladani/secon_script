import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import random
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

APP_TITLE = "Soil & Groundwater Chemical Test Generator v2"

# -----------------------------
# Chemistry model
# -----------------------------
SOIL_RANGES = {
    "Normal River": {
        "pH": {"Low": (6.5,7.0), "Normal": (7.0,7.8), "High": (7.8,8.3), "Very High": (8.3,8.6)},
        "Chloride": {"Low": (50,150), "Normal": (150,450), "High": (450,900), "Very High": (900,1700)},
        "Sulphate": {"Low": (50,200), "Normal": (200,400), "High": (400,800), "Very High": (800,1600)},
        "Organic Matter": {"Low": (0.2,0.5), "Normal": (0.5,1.2), "High": (1.2,2.0), "Very High": (2.0,3.0)}
    },
    "Desert River": {
        "pH": {"Low": (7.5,8.0), "Normal": (8.0,8.7), "High": (8.7,9.2), "Very High": (9.2,9.6)},
        "Chloride": {"Low": (200,450), "Normal": (450,1350), "High": (1350,3300), "Very High": (3300,5000)},
        "Sulphate": {"Low": (150,450), "Normal": (450,1400), "High": (1400,4500), "Very High": (4500,7000)},
        "Organic Matter": {"Low": (0.1,0.3), "Normal": (0.3,0.8), "High": (0.8,1.3), "Very High": (1.3,2.0)}
    },
    "Natural Ground / Bare Land": {
        "pH": {"Low": (6.3,6.8), "Normal": (6.8,7.8), "High": (7.8,8.5), "Very High": (8.5,9.0)},
        "Chloride": {"Low": (30,120), "Normal": (120,450), "High": (450,900), "Very High": (900,1800)},
        "Sulphate": {"Low": (30,150), "Normal": (150,450), "High": (400,800), "Very High": (800,1600)},
        "Organic Matter": {"Low": (0.1,0.4), "Normal": (0.4,1.2), "High": (1.2,2.2), "Very High": (2.2,3.5)}
    },
    "Agricultural Field": {
        "pH": {"Low": (6.2,6.8), "Normal": (6.8,7.8), "High": (7.8,8.4), "Very High": (8.4,8.8)},
        "Chloride": {"Low": (50,200), "Normal": (200,450), "High": (450,900), "Very High": (900,1500)},
        "Sulphate": {"Low": (50,200), "Normal": (200,450), "High": (450,1000), "Very High": (1000,1700)},
        "Organic Matter": {"Low": (0.5,1.0), "Normal": (1.0,2.5), "High": (2.5,4.0), "Very High": (4.0,6.0)}
    }
}
GW_RANGES = {
    "Normal River": {
        "pH": {"Low": (6.8,7.1), "Normal": (7.1,7.8), "High": (7.8,8.2), "Very High": (8.2,8.5)},
        "Chloride": {"Low": (20,75), "Normal": (75,250), "High": (250,500), "Very High": (500,1000)},
        "Sulphate": {"Low": (20,75), "Normal": (75,200), "High": (200,400), "Very High": (400,700)}
    },
    "Desert River": {
        "pH": {"Low": (7.2,7.5), "Normal": (7.5,8.1), "High": (8.1,8.5), "Very High": (8.5,8.8)},
        "Chloride": {"Low": (100,250), "Normal": (250,700), "High": (700,1500), "Very High": (1500,3000)},
        "Sulphate": {"Low": (100,250), "Normal": (250,600), "High": (600,1200), "Very High": (1200,2000)}
    },
    "Natural Ground / Bare Land": {
        "pH": {"Low": (6.7,7.1), "Normal": (7.1,7.8), "High": (7.8,8.3), "Very High": (8.3,8.6)},
        "Chloride": {"Low": (15,60), "Normal": (60,220), "High": (220,450), "Very High": (450,900)},
        "Sulphate": {"Low": (15,60), "Normal": (60,200), "High": (200,400), "Very High": (400,700)}
    },
    "Agricultural Field": {
        "pH": {"Low": (6.5,6.9), "Normal": (6.9,7.7), "High": (7.7,8.2), "Very High": (8.2,8.5)},
        "Chloride": {"Low": (20,100), "Normal": (100,300), "High": (300,600), "Very High": (600,1000)},
        "Sulphate": {"Low": (20,75), "Normal": (75,250), "High": (250,500), "Very High": (500,900)}
    }
}
CONDITION_OPTIONS = ["Dry", "Slightly Moist", "Moist", "Wet", "Saturated", "Saline", "Highly Saline"]
SALINITY_CONTROL_OPTIONS = ["Low", "Normal", "High", "Very High"]
GW_CONDITION_OPTIONS = ["Normal", "Slightly Saline", "Saline", "Highly Saline"]
GW_SALINITY_CONTROL_OPTIONS = ["Low", "Normal", "High", "Very High"]

CONDITION_FACTOR = {
    "Dry": (0.90, 0.90, 0.90),
    "Slightly Moist": (0.97, 0.97, 0.98),
    "Moist": (1.00, 1.00, 1.00),
    "Wet": (1.03, 1.05, 1.02),
    "Saturated": (1.05, 1.08, 1.03),
    "Saline": (1.15, 1.25, 0.95),
    "Highly Saline": (1.30, 1.45, 0.90)
}

GW_CONDITION_FACTOR = {
    "Normal": (1.00, 1.00),
    "Slightly Saline": (1.15, 1.15),
    "Saline": (1.40, 1.35),
    "Highly Saline": (1.80, 1.70)
}

DOMINANCE_PROB = {
    "Normal River": [("Chloride dominant", 40), ("Mixed", 45), ("Sulphate dominant", 15)],
    "Desert River": [("Chloride dominant", 35), ("Mixed", 40), ("Sulphate dominant", 25)],
    "Natural Ground / Bare Land": [("Chloride dominant", 35), ("Mixed", 50), ("Sulphate dominant", 15)],
    "Agricultural Field": [("Chloride dominant", 35), ("Mixed", 50), ("Sulphate dominant", 15)]
}

DOMINANCE_RATIO = {
    # Ratio = Chloride / Sulphate.
    "Chloride dominant": (1.30, 2.40),
    "Mixed": (0.90, 1.40),
    "Sulphate dominant": (0.55, 0.95)
}

def choose_dominance(env, residue):
    if residue == "Strong":
        probs = [("Chloride dominant", 18), ("Mixed", 27), ("Sulphate dominant", 55)]
    elif residue == "Visible":
        probs = [("Chloride dominant", 22), ("Mixed", 33), ("Sulphate dominant", 45)]
    elif residue == "Slight":
        probs = [("Chloride dominant", 28), ("Mixed", 42), ("Sulphate dominant", 30)]
    else:
        probs = DOMINANCE_PROB[env]
    return wchoice(probs)

SULPHATE_SCALE = {
    "Normal River": 0.78,
    "Desert River": 0.72,
    "Natural Ground / Bare Land": 0.78,
    "Agricultural Field": 0.76
}

GROUNDWATER_SULPHATE_SCALE = {
    "Normal River": 0.80,
    "Desert River": 0.75,
    "Natural Ground / Bare Land": 0.80,
    "Agricultural Field": 0.78
}

def apply_dominance(cl, so4, dominance, variation=0.10):
    lo, hi = DOMINANCE_RATIO[dominance]
    ratio = random.uniform(lo, hi) * random.uniform(1-variation, 1+variation)
    magnitude = max(1.0, (cl + so4) / 2.0)
    cl_new = magnitude * 2.0 * ratio / (1.0 + ratio)
    so_new = magnitude * 2.0 / (1.0 + ratio)
    return max(1, int(round(cl_new))), max(1, int(round(so_new)))


SOIL_EFFECT = {
    "SM": (0.85,1.00,0.75), "SC": (0.95,1.10,0.90),
    "ML": (0.90,1.05,1.05), "CL": (0.95,1.10,1.15),
    "CI": (1.00,1.15,1.20), "CH": (1.00,1.20,1.30),
    "SP": (0.80,0.95,0.65), "SW": (0.80,1.00,0.70),
    "GP": (0.75,0.95,0.55), "GW": (0.75,0.95,0.60),
    "GM": (0.85,1.00,0.80), "GC": (0.90,1.05,0.90),
    "SILT": (0.90,1.05,1.00), "CLAY": (1.00,1.15,1.20)
}

COND_PROB = [("Low",20),("Normal",55),("High",20),("Very High",5)]
DESERT_PROB = [("Low",10),("Normal",45),("High",30),("Very High",15)]

def wchoice(items):
    return random.choices([x[0] for x in items], weights=[x[1] for x in items], k=1)[0]

def bounded_gauss(a,b):
    x=random.gauss((a+b)/2,(b-a)/4.0)
    return max(a,min(b,x))

def salt_condition(env, residue):
    if env == "Desert River":
        c=wchoice(DESERT_PROB)
    else:
        c=wchoice(COND_PROB)
    if residue == "Strong":
        c=wchoice([("Normal",15),("High",50),("Very High",35)])
    elif residue == "Visible":
        c=wchoice([("Normal",35),("High",50),("Very High",15)])
    elif residue == "Slight":
        c=wchoice([("Low",20),("Normal",55),("High",25)])
    return c


def depth_zone(depth):
    if depth <= 2.0:
        return "Surface"
    if depth <= 5.0:
        return "Intermediate"
    return "Deep"

def depth_factor(depth, gw_depth, env):
    """Returns multipliers for chloride, sulphate and organic matter."""
    zone = depth_zone(depth)
    # Surface evaporation/accumulation in arid areas; otherwise mild surface OM effect.
    if env == "Desert River":
        if zone == "Surface":
            cl, so, om = 1.08, 1.04, 1.08
        elif zone == "Intermediate":
            cl, so, om = 1.00, 1.00, 0.95
        else:
            cl, so, om = 0.98, 0.94, 0.88
    elif env == "Natural Ground / Bare Land":
        if zone == "Surface":
            cl, so, om = 0.98, 1.00, 1.08
        elif zone == "Intermediate":
            cl, so, om = 1.00, 0.99, 0.96
        else:
            cl, so, om = 1.01, 0.95, 0.88
    elif env == "Agricultural Field":
        if zone == "Surface":
            cl, so, om = 1.03, 1.02, 1.15
        elif zone == "Intermediate":
            cl, so, om = 1.00, 0.99, 0.98
        else:
            cl, so, om = 1.02, 0.96, 0.88
    else:
        if zone == "Surface":
            cl, so, om = 0.96, 0.97, 1.08
        elif zone == "Intermediate":
            cl, so, om = 1.00, 1.00, 0.96
        else:
            cl, so, om = 1.02, 1.04, 0.88

    # Groundwater influence depends on the sample's position relative to
    # groundwater. At or below groundwater depth, the soil is treated as
    # being in/near the saturated zone.
    if depth >= gw_depth:
        gw_influence = 0.30
    else:
        gap = gw_depth - depth
        if gap <= 1.5:
            gw_influence = 0.25
        elif gap <= 3.0:
            gw_influence = 0.12
        elif gap <= 5.0:
            gw_influence = 0.05
        else:
            gw_influence = 0.0

    return cl, so, om, gw_influence, zone

def correlated_natural_ground_ph():
    # Natural/bare ground is generally near-neutral to mildly alkaline.
    # A borehole should have a coherent pH signature rather than unrelated
    # pH values at S-1 and S-2.
    return bounded_gauss(7.05, 7.75)

def soil_result(env, soil_type, residue, bore_profile, sample_index, condition, depth, gw_depth, dominance, salinity_control, ph_anchor=None):
    r=SOIL_RANGES[env]
    # Salinity is now an explicit user input for each soil depth/sample.
    # Depth may influence concentration modestly, but it no longer selects
    # the salinity class.
    cond=salinity_control
    pcond = cond if env=="Desert River" else wchoice([(cond,70),("Normal",20),("High",10)])
    if env == "Natural Ground / Bare Land" and ph_anchor is not None:
        # Keep S-1 and S-2 correlated within the same natural-ground horizon,
        # with only small sample-to-sample variation.
        pH = max(r["pH"]["Low"][0], min(r["pH"]["Very High"][1],
                 ph_anchor + random.uniform(-0.10, 0.10)))
    else:
        pH=bounded_gauss(*r["pH"][pcond])
    cl=bounded_gauss(*r["Chloride"][cond])
    so=bounded_gauss(*r["Sulphate"][wchoice([(cond,70),("Normal",15),("High",15)])])
    omcond=wchoice(COND_PROB)
    om=bounded_gauss(*r["Organic Matter"][omcond])
    eff=SOIL_EFFECT.get(soil_type,(1,1,1))
    cf=CONDITION_FACTOR.get(condition,(1,1,1))
    df_cl, df_so, df_om, gw_inf, zone = depth_factor(depth, gw_depth, env)
    cl*=eff[0]*cf[0]*df_cl
    so*=eff[1]*cf[1]*df_so
    om*=eff[2]*cf[2]*df_om

    # Near groundwater, gently pull soil salts toward the borehole groundwater profile.
    # This is deliberately a modest effect so soil and groundwater remain distinct tests.
    if gw_inf > 0:
        gw_cl_proxy = {
            "Normal River": (180, 130),
            "Desert River": (650, 500),
            "Natural Ground / Bare Land": (150, 120),
            "Agricultural Field": (250, 180)
        }
        pcl, pso = gw_cl_proxy[env]
        cl = cl*(1-gw_inf) + pcl*gw_inf
        so = so*(1-gw_inf) + (pso*0.85 + so*0.15)*gw_inf

    if condition in ("Saline", "Highly Saline"):
        pH += random.uniform(0.03, 0.15)
    # Residue input constrains soluble salt levels
    if residue=="None":
        cl*=random.uniform(0.75,0.95); so*=random.uniform(0.75,0.95)
    elif residue=="Slight":
        cl*=random.uniform(0.95,1.10); so*=random.uniform(0.95,1.10)
    elif residue=="Visible":
        cl*=random.uniform(1.10,1.30); so*=random.uniform(1.10,1.35)
    else:
        cl*=random.uniform(1.25,1.60); so*=random.uniform(1.25,1.60)
    if env=="Desert River" and max(cl,so)>0:
        pH=min(9.6,pH+random.uniform(0,0.18)*(max(cl/3500,so/6000)))
    so *= SULPHATE_SCALE.get(env, 0.78)
    cl, so = apply_dominance(cl, so, dominance, 0.12)
    return round(pH,2), cl, so, round(min(6,om),2), cond, zone, dominance

def groundwater_result(env, bore_profile, condition, dominance, salinity_control):
    r=GW_RANGES[env]
    # Groundwater salinity class is exactly the user's input.
    cond=salinity_control
    pcond=cond
    ph=bounded_gauss(*r["pH"][pcond])
    cl=bounded_gauss(*r["Chloride"][cond])
    so=bounded_gauss(*r["Sulphate"][cond])
    gcf=GW_CONDITION_FACTOR.get(condition,(1.0,1.0))
    cl*=gcf[0]; so*=gcf[1]
    if condition in ("Saline", "Highly Saline"):
        ph += random.uniform(0.02, 0.10)
    if env=="Desert River":
        ph=min(8.8,ph+random.uniform(0,0.08))
    so *= GROUNDWATER_SULPHATE_SCALE.get(env, 0.80)
    cl, so = apply_dominance(cl, so, dominance, 0.15)
    return round(ph,2),max(1,int(round(cl))),max(1,int(round(so))),cond,dominance

class App:
    def __init__(self,root):
        self.root=root
        self.root.title(APP_TITLE)
        self.root.geometry("1450x850")
        self.root.minsize(1200,700)
        self.bore_rows=[]
        self.soil_results=[]
        self.gw_results=[]
        self.seed_var=tk.StringVar(value=str(random.randint(100000,999999)))
        self.count_var=tk.IntVar(value=5)
        self.env_var=tk.StringVar(value="Normal River")
        self.project_var=tk.StringVar()
        self.location_var=tk.StringVar()
        self.build()

    def build(self):
        top=ttk.Frame(self.root,padding=12); top.pack(fill="x")
        ttk.Label(top,text=APP_TITLE,font=("Segoe UI",18,"bold")).pack(side="left")
        ttk.Label(top,text="Synthetic data only",font=("Segoe UI",9)).pack(side="right")

        f=ttk.LabelFrame(self.root,text="Project & Generation Settings",padding=10)
        f.pack(fill="x",padx=12,pady=5)
        ttk.Label(f,text="Project:").grid(row=0,column=0,sticky="w")
        ttk.Entry(f,textvariable=self.project_var,width=28).grid(row=0,column=1,padx=5)
        ttk.Label(f,text="Location:").grid(row=0,column=2,sticky="w",padx=(20,0))
        ttk.Entry(f,textvariable=self.location_var,width=28).grid(row=0,column=3,padx=5)
        ttk.Label(f,text="No. of Boreholes:").grid(row=0,column=4,padx=(20,0))
        tk.Spinbox(f,from_=1,to=500,textvariable=self.count_var,width=8,command=self.make_borehole_rows).grid(row=0,column=5,padx=5)
        ttk.Label(f,text="Environment:").grid(row=0,column=6,padx=(20,0))
        ttk.Combobox(f,textvariable=self.env_var,values=list(SOIL_RANGES),state="readonly",width=22).grid(row=0,column=7,padx=5)
        ttk.Label(f,text="Seed:").grid(row=0,column=8,padx=(20,0))
        ttk.Entry(f,textvariable=self.seed_var,width=12).grid(row=0,column=9,padx=5)
        ttk.Button(f,text="Generate",command=self.generate).grid(row=0,column=10,padx=8)
        ttk.Button(f,text="Import Excel",command=self.import_excel).grid(row=0,column=11,padx=5)
        ttk.Button(f,text="Export Excel",command=self.export).grid(row=0,column=12,padx=5)

        note=ttk.Label(self.root,text="Each borehole: exactly 2 soil chemical samples + exactly 1 groundwater chemical sample. Enter manual depths, soil classification, conditions and white-residue observation. S-2 may be at or below groundwater depth.",padding=(12,5))
        note.pack(fill="x")

        nb=ttk.Notebook(self.root); nb.pack(fill="both",expand=True,padx=12,pady=5)
        self.setup_tab=ttk.Frame(nb); nb.add(self.setup_tab,text="1. Borehole Input")
        self.soil_tab=ttk.Frame(nb); nb.add(self.soil_tab,text="2. Soil Results")
        self.gw_tab=ttk.Frame(nb); nb.add(self.gw_tab,text="3. Groundwater Results")
        self.sum_tab=ttk.Frame(nb); nb.add(self.sum_tab,text="4. Borehole Summary")
        self.method_tab=ttk.Frame(nb); nb.add(self.method_tab,text="5. Test Methods")

        self.make_borehole_editor()
        self.make_result_table(self.soil_tab,"soil")
        self.make_result_table(self.gw_tab,"gw")
        self.make_result_table(self.sum_tab,"summary")
        self.make_methods()

        self.status=tk.StringVar(value="Ready.")
        ttk.Label(self.root,textvariable=self.status,relief="sunken",anchor="w",padding=5).pack(fill="x")

    def make_borehole_editor(self):
        for w in self.setup_tab.winfo_children(): w.destroy()
        wrap=ttk.Frame(self.setup_tab); wrap.pack(fill="both",expand=True)
        canvas=tk.Canvas(wrap,highlightthickness=0)
        vs=ttk.Scrollbar(wrap,orient="vertical",command=canvas.yview)
        inner=ttk.Frame(canvas)
        inner.bind("<Configure>",lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0),window=inner,anchor="nw")
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side="left",fill="both",expand=True); vs.pack(side="right",fill="y")
        headers=["BH ID","S-1 Depth (m)","S-1 Type","S-1 Condition","S-1 Salinity Control","S-2 Depth (m)","S-2 Type","S-2 Condition","S-2 Salinity Control","GW Depth (m)","GW Condition","GW Salinity Control","White Residue"]
        for j,h in enumerate(headers): ttk.Label(inner,text=h,font=("Segoe UI",9,"bold")).grid(row=0,column=j,padx=6,pady=6)
        self.bore_rows=[]
        for i in range(self.count_var.get()):
            vals=[tk.StringVar(value=f"BH-{i+1:02d}"),tk.StringVar(value="1.5"),tk.StringVar(value="CL"),tk.StringVar(value="Moist"),tk.StringVar(value="Normal"),tk.StringVar(value="4.5"),tk.StringVar(value="CL"),tk.StringVar(value="Moist"),tk.StringVar(value="Normal"),tk.StringVar(value="8.0"),tk.StringVar(value="Normal"),tk.StringVar(value="Normal"),tk.StringVar(value="None")]
            ttk.Entry(inner,textvariable=vals[0],width=14).grid(row=i+1,column=0,padx=5,pady=3)
            ttk.Entry(inner,textvariable=vals[1],width=13).grid(row=i+1,column=1)
            ttk.Combobox(inner,textvariable=vals[2],values=list(SOIL_EFFECT),state="readonly",width=11).grid(row=i+1,column=2)
            ttk.Combobox(inner,textvariable=vals[3],values=CONDITION_OPTIONS,state="readonly",width=16).grid(row=i+1,column=3)
            ttk.Combobox(inner,textvariable=vals[4],values=SALINITY_CONTROL_OPTIONS,state="readonly",width=17).grid(row=i+1,column=4)
            ttk.Entry(inner,textvariable=vals[5],width=13).grid(row=i+1,column=5)
            ttk.Combobox(inner,textvariable=vals[6],values=list(SOIL_EFFECT),state="readonly",width=11).grid(row=i+1,column=6)
            ttk.Combobox(inner,textvariable=vals[7],values=CONDITION_OPTIONS,state="readonly",width=16).grid(row=i+1,column=7)
            ttk.Combobox(inner,textvariable=vals[8],values=SALINITY_CONTROL_OPTIONS,state="readonly",width=17).grid(row=i+1,column=8)
            ttk.Entry(inner,textvariable=vals[9],width=13).grid(row=i+1,column=9)
            ttk.Combobox(inner,textvariable=vals[10],values=GW_CONDITION_OPTIONS,state="readonly",width=17).grid(row=i+1,column=10)
            ttk.Combobox(inner,textvariable=vals[11],values=GW_SALINITY_CONTROL_OPTIONS,state="readonly",width=17).grid(row=i+1,column=11)
            ttk.Combobox(inner,textvariable=vals[12],values=["None","Slight","Visible","Strong"],state="readonly",width=13).grid(row=i+1,column=12)
            self.bore_rows.append(vals)

    def make_borehole_rows(self):
        try: n=max(1,min(500,int(self.count_var.get())))
        except: n=5; self.count_var.set(5)
        self.count_var.set(n); self.make_borehole_editor()

    def make_result_table(self,parent,kind):
        frame=ttk.Frame(parent,padding=6); frame.pack(fill="both",expand=True)
        if kind=="soil":
            cols=["BH ID","Sample","Depth (m)","Depth Zone","Soil Type","Condition","Salinity Control","Environment","pH","Chloride (mg/kg)","Sulphate (mg/kg)","Organic Matter (%)","White Residue","Chemical Dominance","Validation"]
        elif kind=="gw":
            cols=["BH ID","Sample","Depth (m)","Environment","Condition","pH","Chloride (mg/L)","Sulphate (mg/L)","Salinity Condition","Chemical Dominance","Validation"]
        else:
            cols=["BH ID","Environment","S-1 Depth","S-1 Zone","S-1 Type","S-1 Condition","S-1 Salinity Control","S-2 Depth","S-2 Zone","S-2 Type","S-2 Condition","S-2 Salinity Control","GW Depth","GW Condition","GW Salinity Control","White Residue","Chemical Dominance","S-1 Cl","S-2 Cl","GW Cl","S-1 SO4","S-2 SO4","GW SO4","Overall Check"]
        tree=ttk.Treeview(frame,columns=cols,show="headings")
        vs=ttk.Scrollbar(frame,orient="vertical",command=tree.yview); hs=ttk.Scrollbar(frame,orient="horizontal",command=tree.xview)
        tree.configure(yscrollcommand=vs.set,xscrollcommand=hs.set)
        tree.pack(side="left",fill="both",expand=True); vs.pack(side="right",fill="y"); hs.pack(side="bottom",fill="x")
        for c in cols:
            tree.heading(c,text=c); tree.column(c,width=max(95,min(170,len(c)*9+20)),anchor="center")
        setattr(self,kind+"_tree",tree)

    def make_methods(self):
        cols=["Parameter","Sample","IS Code / Standard","Method","Unit"]
        tree=ttk.Treeview(self.method_tab,columns=cols,show="headings")
        tree.pack(fill="both",expand=True,padx=8,pady=8)
        rows=[
            ("pH","Soil","IS 2720 (Part 26)","Electrometric / glass electrode","pH"),
            ("Chloride","Soil","IS 2720 (Part 27)","Water extract; argentometric titration","mg/kg"),
            ("Sulphate","Soil","IS 2720 (Part 27)","Water extract; gravimetric / turbidimetric","mg/kg"),
            ("Organic Matter","Soil","IS 2720 (Part 22)","Wet oxidation","%"),
            ("pH","Groundwater","IS 3025 (Part 11)","Electrometric / glass electrode","pH"),
            ("Chloride","Groundwater","IS 3025 (Part 32)","Argentometric titration","mg/L"),
            ("Sulphate","Groundwater","IS 3025 (Part 24)","Turbidimetric / gravimetric","mg/L"),
        ]
        for c in cols: tree.heading(c,text=c); tree.column(c,width=220,anchor="center")
        for r in rows: tree.insert("", "end",values=r)

    def clear_tree(self,t):
        for x in t.get_children(): t.delete(x)

    def validate_borehole(self,vals):
        errors=[]
        if not vals[0].get().strip():
            errors.append("BH ID is required")
        try:
            # Current Borehole Input order:
            # 0=BH ID, 1=S-1 Depth, 2=S-1 Type, 3=S-1 Condition,
            # 4=S-1 Salinity Control, 5=S-2 Depth, 6=S-2 Type,
            # 7=S-2 Condition, 8=S-2 Salinity Control, 9=GW Depth,
            # 10=GW Condition, 11=White Residue
            d1=float(vals[1].get())
            d2=float(vals[5].get())
            dg=float(vals[9].get())

            if d1<=0 or d2<=0 or dg<=0:
                errors.append("depth must be > 0")
            if vals[4].get() not in SALINITY_CONTROL_OPTIONS:
                errors.append("invalid S-1 salinity control")
            if vals[8].get() not in SALINITY_CONTROL_OPTIONS:
                errors.append("invalid S-2 salinity control")
            if vals[11].get() not in GW_SALINITY_CONTROL_OPTIONS:
                errors.append("invalid groundwater salinity control")
            if d1>=d2:
                errors.append("S-1 depth must be less than S-2 depth")
            # Groundwater depth is not a maximum soil depth.
            # S-2 may be at or below the groundwater depth because soil
            # samples can be collected from the saturated zone.
        except (ValueError, TypeError):
            errors.append("invalid depth")
        return errors

    def generate(self):
        try: seed=int(self.seed_var.get())
        except:
            messagebox.showerror("Invalid Seed","Seed must be an integer."); return
        if seed<0: messagebox.showerror("Invalid Seed","Seed must be non-negative."); return
        random.seed(seed)
        env=self.env_var.get()
        self.soil_results=[]; self.gw_results=[]
        self.clear_tree(self.soil_tree); self.clear_tree(self.gw_tree); self.clear_tree(self.summary_tree)
        for vals in self.bore_rows:
            errs=self.validate_borehole(vals)
            if errs:
                messagebox.showerror("Input Validation",f"{vals[0].get()}: "+", ".join(errs)); return
        for vals in self.bore_rows:
            bh=vals[0].get(); d1=float(vals[1].get()); t1=vals[2].get(); c1=vals[3].get(); sc1=vals[4].get(); d2=float(vals[5].get()); t2=vals[6].get(); c2=vals[7].get(); sc2=vals[8].get(); dg=float(vals[9].get()); gwc=vals[10].get(); gwsc=vals[11].get(); residue=vals[12].get()
            profile=salt_condition(env,residue)
            dominance=choose_dominance(env,residue)
            ph_anchor = correlated_natural_ground_ph() if env=="Natural Ground / Bare Land" else None
            s1=soil_result(env,t1,residue,profile,1,c1,d1,dg,dominance,sc1,ph_anchor)
            s2=soil_result(env,t2,residue,profile,2,c2,d2,dg,dominance,sc2,ph_anchor)
            gw=groundwater_result(env,profile,gwc,dominance,gwsc)
            for idx,(depth,typ,res) in enumerate([(d1,t1,s1),(d2,t2,s2)],1):
                p,cl,so,om,cond,zone,dom=res
                valid=self.validate_result(env,cl,so,p,residue)
                row=[bh,f"S-{idx}",depth,zone,typ,(c1 if idx==1 else c2),cond,env,p,cl,so,om,residue,dom,valid]
                self.soil_tree.insert("", "end",values=row); self.soil_results.append(row)
            validgw=self.validate_gw(env,gw[1],gw[2],gw[0])
            grow=[bh,"GW-1",dg,env,gwc,gw[0],gw[1],gw[2],gw[3],gw[4],validgw]
            self.gw_tree.insert("", "end",values=grow); self.gw_results.append(grow)
            self.summary_tree.insert("", "end",values=[bh,env,d1,depth_zone(d1),t1,c1,sc1,d2,depth_zone(d2),t2,c2,sc2,dg,gwc,gwsc,residue,dominance,s1[1],s2[1],gw[1],s1[2],s2[2],gw[2],"PASS"])
        self.status.set(f"Generated {len(self.bore_rows)} boreholes: {len(self.soil_results)} soil tests + {len(self.gw_results)} groundwater tests. Seed={seed}")

    def validate_result(self,env,cl,so,ph,residue):
        r=SOIL_RANGES[env]
        if not (r["pH"]["Low"][0] <= ph <= r["pH"]["Very High"][1]): return "CHECK: pH"
        if cl<=0 or so<=0: return "CHECK: salt"
        if residue=="None" and env!="Desert River" and (cl>2200 or so>2700): return "CHECK: residue/salt"
        return "PASS"

    def validate_gw(self,env,cl,so,ph):
        r=GW_RANGES[env]
        if not (r["pH"]["Low"][0] <= ph <= r["pH"]["Very High"][1]): return "CHECK: pH"
        if cl<=0 or so<=0: return "CHECK: salt"
        return "PASS"

    def import_excel(self):
        """Import borehole input data from the supplied Excel template."""
        fn = filedialog.askopenfilename(
            title="Import Borehole Input Excel",
            filetypes=[("Excel Workbook", "*.xlsx *.xlsm")]
        )
        if not fn:
            return

        required = [
            "BH ID", "S-1 Depth (m)", "S-1 Type", "S-1 Condition", "S-1 Salinity Control",
            "S-2 Depth (m)", "S-2 Type", "S-2 Condition", "S-2 Salinity Control",
            "GW Depth (m)", "GW Condition", "GW Salinity Control", "White Residue"
        ]

        try:
            wb = load_workbook(fn, data_only=True)
            if "Borehole Input" in wb.sheetnames:
                ws = wb["Borehole Input"]
            else:
                ws = wb[wb.sheetnames[0]]

            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                raise ValueError("The selected workbook is empty.")

            headers = [str(x).strip() if x is not None else "" for x in rows[0]]
            missing = [h for h in required if h not in headers]
            if missing:
                raise ValueError(
                    "Required column(s) missing:\\n\\n" + "\\n".join(missing)
                )

            idx = {h: headers.index(h) for h in required}
            imported = []

            for excel_row, row in enumerate(rows[1:], start=2):
                if not any(v is not None and str(v).strip() != "" for v in row):
                    continue

                vals = []
                for h in required:
                    value = row[idx[h]] if idx[h] < len(row) else None
                    if value is None:
                        value = ""
                    vals.append(str(value).strip())

                # Skip purely instructional/example rows.
                bh_id = vals[0]
                if not bh_id or bh_id.startswith("#"):
                    continue

                imported.append(vals)

            if not imported:
                raise ValueError("No valid borehole rows were found.")

            # Resize borehole count.
            self.count_var.set(len(imported))
            self.make_borehole_editor()

            for gui_vals, imported_vals in zip(self.bore_rows, imported):
                for target, value in zip(gui_vals, imported_vals):
                    target.set(value)

            # Optional Settings sheet
            if "Generation Settings" in wb.sheetnames:
                sws = wb["Generation Settings"]
                settings = {}
                for r in sws.iter_rows(min_row=1, max_col=2, values_only=True):
                    if r[0] is not None:
                        settings[str(r[0]).strip()] = "" if r[1] is None else str(r[1]).strip()

                if settings.get("Project") is not None:
                    self.project_var.set(settings.get("Project", ""))
                if settings.get("Location") is not None:
                    self.location_var.set(settings.get("Location", ""))
                if settings.get("Environment") in SOIL_RANGES:
                    self.env_var.set(settings["Environment"])
                if settings.get("Seed", "").strip():
                    self.seed_var.set(settings["Seed"])

            self.status.set(
                f"Imported {len(imported)} borehole(s) from Excel. "
                f"Review inputs, then click Generate."
            )
            messagebox.showinfo(
                "Import Complete",
                f"Successfully imported {len(imported)} borehole(s).\\n\\n"
                "Please review the Borehole Input tab and click Generate."
            )

        except Exception as e:
            messagebox.showerror("Excel Import Error", str(e))

    def export(self):
        if not self.soil_results:
            messagebox.showwarning("No Data","Generate data first."); return
        fn=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel Workbook","*.xlsx")],title="Save Chemical Generator Results")
        if not fn:return
        wb=Workbook(); ws=wb.active; ws.title="Soil Chemical Results"
        sheets=[("Soil Chemical Results",self.soil_results),("Groundwater Chemical Results",self.gw_results)]
        summary=[]
        for item_id in self.summary_tree.get_children():
            summary.append(list(self.summary_tree.item(item_id)["values"]))
        sheets.append(("Borehole Summary",summary))
        methods=[
            ["Parameter","Sample","IS Code / Standard","Method","Unit"],
            ["pH","Soil","IS 2720 (Part 26)","Electrometric / glass electrode","pH"],
            ["Chloride","Soil","IS 2720 (Part 27)","Water extract; argentometric titration","mg/kg"],
            ["Sulphate","Soil","IS 2720 (Part 27)","Water extract; gravimetric / turbidimetric","mg/kg"],
            ["Organic Matter","Soil","IS 2720 (Part 22)","Wet oxidation","%"],
            ["pH","Groundwater","IS 3025 (Part 11)","Electrometric / glass electrode","pH"],
            ["Chloride","Groundwater","IS 3025 (Part 32)","Argentometric titration","mg/L"],
            ["Sulphate","Groundwater","IS 3025 (Part 24)","Turbidimetric / gravimetric","mg/L"],
        ]
        sheets.append(("Test Methods",methods))
        settings=[
            ["Project",self.project_var.get()],["Location",self.location_var.get()],
            ["Environment",self.env_var.get()],["Boreholes",self.count_var.get()],
            ["Soil samples / borehole",2],["Groundwater samples / borehole",1],
            ["Seed",self.seed_var.get()],["Depth Model","Surface / Intermediate / Deep zones with groundwater proximity influence"],["Import Template","Use the Borehole Input worksheet for direct import."],["Note","Synthetic data for testing/workflow purposes; not actual laboratory measurements."]
        ]
        sheets.append(("Generation Settings",settings))
        for idx,(name,data) in enumerate(sheets):
            if idx==0: sh=ws
            else: sh=wb.create_sheet(name)
            if name=="Soil Chemical Results":
                headers=["BH ID","Sample","Depth (m)","Depth Zone","Soil Type","Condition","Salinity Control","Environment","pH","Chloride (mg/kg)","Sulphate (mg/kg)","Organic Matter (%)","White Residue","Chemical Dominance","Validation"]
                data=[headers]+data
            elif name=="Groundwater Chemical Results":
                headers=["BH ID","Sample","Depth (m)","Environment","Condition","pH","Chloride (mg/L)","Sulphate (mg/L)","Salinity Condition","Chemical Dominance","Validation"]
                data=[headers]+data
            elif name=="Borehole Summary":
                data=[["BH ID","Environment","S-1 Depth","S-1 Zone","S-1 Type","S-1 Condition","S-2 Depth","S-2 Zone","S-2 Type","S-2 Condition","GW Depth","GW Condition","White Residue","Chemical Dominance","S-1 Cl","S-2 Cl","GW Cl","S-1 SO4","S-2 SO4","GW SO4","Overall Check"]]+data
            for r,row in enumerate(data,1):
                for c,val in enumerate(row,1):
                    cell=sh.cell(r,c,val); cell.alignment=Alignment(horizontal="center",vertical="center")
                    if r==1: cell.font=Font(bold=True)
            sh.freeze_panes="A2"
            sh.auto_filter.ref=sh.dimensions
            for col in range(1,sh.max_column+1):
                vals=[str(sh.cell(r,col).value or "") for r in range(1,min(sh.max_row,100)+1)]
                sh.column_dimensions[get_column_letter(col)].width=min(35,max(12,max(map(len,vals))+2))
        try:
            wb.save(fn)
            messagebox.showinfo("Export Complete",f"Workbook saved successfully:\n{fn}")
        except Exception as e:
            messagebox.showerror("Export Error",str(e))

if __name__=="__main__":
    root=tk.Tk()
    App(root)
    root.mainloop()
