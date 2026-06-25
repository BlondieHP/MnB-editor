#!/usr/bin/env python
"""
MnBEditor GUI — Python/tkinter replica of the original VB6 editor.
Layout matches frmMain + frmItems/frmTroops/frmFactions.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json, os, sys

# === Constants from game ===
ITP_TYPES = {1:"horse",2:"1h_wpn",3:"2h_wpn",4:"polearm",5:"arrows",6:"bolts",
    7:"shield",8:"bow",9:"crossbow",10:"thrown",11:"goods",12:"head_arm",
    13:"body_arm",14:"foot_arm",15:"hand_arm",16:"pistol",17:"musket",
    18:"bullets",19:"animal",20:"book"}

ITEM_FLAGS = ["Unique","Always Loot","No Parry","Default Ammo","Merchandise",
    "Wooden Atk","Wooden Parry","Food","Cant Reload Horse","Two Handed",
    "Primary","Secondary","Covers Legs","Consumable","Bonus vs Shield",
    "Penalty Shield","Cant Use Horse","Civilian","Fit to Head","Covers Head",
    "Crush Through","Knock Back","Remove on Use","Unbalanced","Covers Beard",
    "No Pickup","Can Knock Down"]

# === FileParser (embedded for standalone use) ===
class FileParser:
    def __init__(self, data): self.buf, self.pos, self.n = bytes(data), 0, len(data)
    def _read(self):
        chars=[];
        while self.pos<self.n:
            b=self.buf[self.pos]; self.pos+=1
            if b in(10,13,32):
                if chars: break
            else: chars.append(chr(b))
        return ''.join(chars)
    def word(self): return self._read()
    def long(self):
        try: return int(self._read())
        except: return 0
    def double(self):
        try: return float(self._read())
        except: return 0.0
    def skip(self,n):
        for _ in range(n): self._read()

# === Models ===
class Item:
    def __init__(self):
        self.id=0; self.db=""; self.name=""; self.tex=""
        self.mcnt=0; self.mesh_n=[]; self.mesh_p=[]; self.type=""; self.act=""
        self.price=0; self.prefix=""; self.weight=""; self.abundance=0
        self.ha=0; self.ba=0; self.la=0; self.diff=0; self.hp=0; self.spd=0
        self.ms=0; self.wlen=0; self.ammo=0; self.thrust=0; self.swing=0
        self.fcnt=0; self.facs=[]; self.tcnt=0; self.edited=False
    def __repr__(self): return f"[{self.id}] {self.db}"

class Troop:
    def __init__(self):
        self.id=0; self.sid=""; self.name=""; self.plural=""
        self.uw=""; self.flags=""; self.scene=0; self.reserved=0; self.faction=0
        self.up1=0; self.up2=0; self.inv=[(0,0)]*64
        self.lv=0; self.str=0; self.agi=0; self.int=0; self.cha=0
        self.profs=[0]*7; self.skills=[0]*6; self.face=[""]*8; self.edited=False
    def __repr__(self): return f"[{self.id}] {self.sid}"

class Faction:
    def __init__(self):
        self.id=0; self.sid=""; self.name=""; self.flags=0
        self.color=""; self.rels=[]; self.reserved=0; self.edited=False

class Repo:
    def __init__(self):
        self.items=[]; self.troops=[]; self.factions=[]
        self.pt=[]; self.parties=[]; self.scenes=[]; self.modpath=""; self.modname=""

# === Parsers ===
def load_items(repo, path):
    p=FileParser(open(path,'rb').read())
    p.word();p.word();p.word(); n=p.long()
    for i in range(n):
        it=Item(); it.id=i; it.db=p.word(); it.name=p.word(); it.tex=p.word()
        it.mcnt=p.long()
        for _ in range(it.mcnt): it.mesh_n.append(p.word()); it.mesh_p.append(p.word())
        it.type=p.word(); it.act=p.word(); it.price=p.long(); it.prefix=p.word()
        it.weight=p.word(); it.abundance=p.long(); it.ha=p.long(); it.ba=p.long()
        it.la=p.long(); it.diff=p.long(); it.hp=p.long(); it.spd=p.long()
        it.ms=p.long(); it.wlen=p.long(); it.ammo=p.long(); it.thrust=p.long()
        it.swing=p.long(); it.fcnt=p.long()
        for _ in range(it.fcnt): it.facs.append(p.long())
        it.tcnt=p.long()
        for _ in range(it.tcnt):
            p.double(); ac=p.long()
            for _ in range(ac): p.word(); p.skip(p.long())
        repo.items.append(it)

def load_troops(repo, path):
    p=FileParser(open(path,'rb').read())
    p.word();p.word();p.word(); n=p.long()
    for i in range(n):
        t=Troop(); t.id=i; t.sid=p.word(); t.name=p.word(); t.plural=p.word()
        t.uw=p.word(); t.flags=p.word(); t.scene=p.long(); t.reserved=p.long()
        t.faction=p.long(); t.up1=p.long(); t.up2=p.long()
        t.inv=[(p.long(),p.long()) for _ in range(64)]
        t.lv=int(p.long());t.str=int(p.long());t.agi=int(p.long());t.int=int(p.long());t.cha=int(p.long())
        t.profs=[p.long() for _ in range(7)]
        t.skills=[p.long() for _ in range(6)]
        t.face=[p.word() for _ in range(8)]
        repo.troops.append(t)

def load_factions(repo, path):
    p=FileParser(open(path,'rb').read())
    p.word();p.word();p.word(); n=p.long()
    for i in range(n):
        f=Faction(); f.id=i; f.sid=p.word(); f.name=p.word(); f.flags=p.long()
        f.color=p.word()
        f.rels=[p.double() for _ in range(n)]
        f.reserved=p.long()
        repo.factions.append(f)

def load_mod(repo, path):
    repo.modpath=path; repo.modname=os.path.basename(path)
    for fn, loader in [("item_kinds1.txt",load_items),("troops.txt",load_troops),
                        ("factions.txt",load_factions)]:
        fp=os.path.join(path,fn)
        if os.path.exists(fp): loader(repo, fp)

# ====== GUI ======
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MnBEditor — M&B Warband MOD Editor")
        self.geometry("1100x700")
        self.repo=Repo()
        self._build_menu()
        self._build_ui()

    def _build_menu(self):
        bar=tk.Menu(self)
        fmenu=tk.Menu(bar,tearoff=0)
        fmenu.add_command(label="Open MOD...", command=self._open)
        fmenu.add_command(label="Save All", command=self._save)
        fmenu.add_separator()
        fmenu.add_command(label="Exit", command=self.quit)
        bar.add_cascade(label="File",menu=fmenu)
        emenu=tk.Menu(bar,tearoff=0)
        emenu.add_command(label="Items", command=lambda:self._open_editor("items"))
        emenu.add_command(label="Troops", command=lambda:self._open_editor("troops"))
        emenu.add_command(label="Factions", command=lambda:self._open_editor("factions"))
        bar.add_cascade(label="Edit",menu=emenu)
        self.config(menu=bar)

    def _build_ui(self):
        # Toolbar
        tb=tk.Frame(self)
        tk.Button(tb,text="Open MOD",command=self._open).pack(side=tk.LEFT,padx=2)
        tk.Button(tb,text="Save All",command=self._save).pack(side=tk.LEFT,padx=2)
        tk.Button(tb,text="Validate",command=self._validate).pack(side=tk.LEFT,padx=2)
        self._status=tk.Label(tb,text="Ready — Open a MOD to begin",anchor=tk.W)
        self._status.pack(side=tk.LEFT,padx=20,fill=tk.X,expand=True)
        tb.pack(side=tk.TOP,fill=tk.X)

        # Left: TreeView
        left=tk.Frame(self,width=200,bg='white',relief=tk.SUNKEN,bd=1)
        left.pack(side=tk.LEFT,fill=tk.Y); left.pack_propagate(False)
        self._tree=ttk.Treeview(left,show='tree',height=30)
        self._tree.pack(fill=tk.BOTH,expand=True)
        self._tree.bind('<<TreeviewSelect>>',self._on_tree_select)

        # Right: Container frame for child editors
        self._container=tk.Frame(self,bg='gray85')
        self._container.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)

    def _build_tree(self):
        self._tree.delete(*self._tree.get_children())
        def add(n,txt): return self._tree.insert('','end',text=f"{txt} ({n})",tags=(txt,))
        add(len(self.repo.items),"Items")
        add(len(self.repo.troops),"Troops")
        add(len(self.repo.factions),"Factions")
        add(len(self.repo.parties),"Parties")
        add(len(self.repo.pt),"Party Templates")
        add(len(self.repo.scenes),"Scenes")

    def _on_tree_select(self, ev):
        sel=self._tree.selection()
        if not sel: return
        tag=self._tree.item(sel[0],'tags')[0]
        self._open_editor(tag.lower().replace(' ','_'))

    def _open(self):
        d=filedialog.askdirectory(title="Select MOD directory")
        if not d: return
        self._status.config(text="Loading..."); self.update()
        try:
            load_mod(self.repo, d)
            self._build_tree()
            self._status.config(text=f"Loaded: {len(self.repo.items)} items, {len(self.repo.troops)} troops, {len(self.repo.factions)} factions")
        except Exception as e:
            messagebox.showerror("Error",str(e))

    def _save(self):
        if not self.repo.items: return
        messagebox.showinfo("Save","Save not yet implemented in Python.\nUse the .NET CLI: dotnet run --project MnBEditor.Cli")

    def _validate(self):
        if not self.repo.items: return
        issues=[]
        for i,t in enumerate(self.repo.troops):
            if t.faction>=len(self.repo.factions): issues.append(f"Troop[{i}] bad faction {t.faction}")
        msg=f"{len(issues)} issues found" if issues else "Validation PASSED"
        messagebox.showinfo("Validate",msg)

    def _open_editor(self, etype):
        # Clear container
        for w in self._container.winfo_children(): w.destroy()
        if etype in('items','item'):
            ItemEditor(self._container, self.repo)
        elif etype in('troops','troop'):
            TroopEditor(self._container, self.repo)
        elif etype in('factions','faction'):
            FactionEditor(self._container, self.repo)
        else:
            tk.Label(self._container,text=f"Editor for {etype}\nnot yet implemented",
                    font=('',14)).place(relx=0.5,rely=0.5,anchor=tk.CENTER)

# ====== ITEM EDITOR ======
class ItemEditor(tk.Frame):
    def __init__(self, parent, repo):
        super().__init__(parent, bg='white'); self.repo=repo; self.items=repo.items
        self.pack(fill=tk.BOTH,expand=True); self._sel=-1; self._build()

    def _build(self):
        # Left panel: Listbox + search
        left=tk.Frame(self,width=380,bg='white')
        left.pack(side=tk.LEFT,fill=tk.Y); left.pack_propagate(False)

        sf=tk.Frame(left,bg='white'); sf.pack(fill=tk.X,padx=4,pady=2)
        self._search_var=tk.StringVar()
        tk.Entry(sf,textvariable=self._search_var,width=28).pack(side=tk.LEFT)
        tk.Button(sf,text="Find",command=self._filter,width=5).pack(side=tk.LEFT,padx=2)

        self._lb=tk.Listbox(left,width=42,font=('Consolas',9),bg='white')
        self._lb.pack(fill=tk.BOTH,expand=True,padx=4,pady=2)
        self._lb.bind('<<ListboxSelect>>',self._on_select)
        self._fill_list()

        # Right panel: Notebook
        nb=ttk.Notebook(self); nb.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        self._build_basic_tab(nb)
        self._build_mod_tab(nb)

        # Bottom buttons
        bf=tk.Frame(self,bg='#f0f0f0',height=36); bf.pack(side=tk.BOTTOM,fill=tk.X)
        self._status_lbl=tk.Label(bf,text="",anchor=tk.W,bg='#f0f0f0')
        self._status_lbl.pack(side=tk.LEFT,padx=8,fill=tk.X,expand=True)
        tk.Button(bf,text="Reset",command=self._reset,width=8).pack(side=tk.RIGHT,padx=4,pady=3)
        tk.Button(bf,text="Apply",command=self._apply,width=8,bg='#ffaaaa').pack(side=tk.RIGHT,padx=4,pady=3)
        tk.Button(bf,text="Output",command=self._output,width=8,bg='#aaffaa').pack(side=tk.RIGHT,padx=4,pady=3)

    def _build_basic_tab(self, nb):
        f=tk.Frame(nb); nb.add(f,text="Basic Properties")
        # Use grid for label+entry pairs
        fields=[("DB Name:","_db"),("Display:","_name"),("Texture:","_tex"),
            ("Price:","_price"),("Weight:","_weight"),("Abundance:","_abun"),
            ("Prefix:","_pfx"),("Mesh Count:","_mcnt"),
            ("Head Armor:","_ha"),("Body Armor:","_ba"),("Leg Armor:","_la"),
            ("Difficulty:","_diff"),("Hit Points:","_hp"),("Speed:","_spd"),
            ("Missile Spd:","_ms"),("Weapon Len:","_wlen"),("Max Ammo:","_ammo"),
            ("Thrust Dmg:","_thrust"),("Swing Dmg:","_swing"),
            ("Type Flags:","_type"),("Action:","_act")]
        self._entries={}
        for i,(lbl,key) in enumerate(fields):
            tk.Label(f,text=lbl,anchor=tk.E,width=14).grid(row=i,column=0,sticky=tk.E,padx=4,pady=1)
            e=tk.Entry(f,width=30); e.grid(row=i,column=1,sticky=tk.W,padx=4,pady=1)
            self._entries[key]=e

        # Flags frame
        tk.Label(f,text="Flags:",anchor=tk.NE,width=14).grid(row=len(fields),column=0,sticky=tk.NE,padx=4)
        ff=tk.Frame(f); ff.grid(row=len(fields),column=1,sticky=tk.W)
        self._flag_vars={}
        for i,fl in enumerate(ITEM_FLAGS):
            v=tk.IntVar(); self._flag_vars[fl]=v
            tk.Checkbutton(ff,text=fl,variable=v,anchor=tk.W,width=18).grid(row=i//3,column=i%3,sticky=tk.W)

    def _build_mod_tab(self, nb):
        f=tk.Frame(nb); nb.add(f,text="Mod Data")
        tk.Label(f,text="Faction Count:").grid(row=0,column=0,sticky=tk.E,padx=4,pady=2)
        self._entries['_fcnt']=tk.Entry(f,width=10)
        self._entries['_fcnt'].grid(row=0,column=1,sticky=tk.W,padx=4)
        tk.Label(f,text="Factions:").grid(row=1,column=0,sticky=tk.NE,padx=4,pady=2)
        self._fac_list=tk.Listbox(f,height=6,width=30); self._fac_list.grid(row=1,column=1,sticky=tk.W)
        tk.Label(f,text="Trigger Count:").grid(row=2,column=0,sticky=tk.E,padx=4,pady=2)
        self._entries['_tcnt']=tk.Entry(f,width=10)
        self._entries['_tcnt'].grid(row=2,column=1,sticky=tk.W,padx=4)

    def _fill_list(self):
        self._lb.delete(0,tk.END)
        for it in self.items:
            self._lb.insert(tk.END,f"{it.id:4d}  {it.db:35s} {it.name:25s} {it.price:6d}")

    def _filter(self):
        q=self._search_var.get().lower()
        self._lb.delete(0,tk.END)
        for it in self.items:
            if not q or q in it.db.lower() or q in it.name.lower():
                self._lb.insert(tk.END,f"{it.id:4d}  {it.db:35s} {it.name:25s} {it.price:6d}")
        self._status_lbl.config(text=f"Filtered: {self._lb.size()}/{len(self.items)}")

    def _on_select(self, ev):
        sel=self._lb.curselection()
        if not sel: return
        idx=int(self._lb.get(sel[0]).split()[0])
        self._sel=idx; it=self.items[idx]
        for k,e in self._entries.items():
            e.delete(0,tk.END)
            if k=='_db': e.insert(0,it.db)
            elif k=='_name': e.insert(0,it.name)
            elif k=='_tex': e.insert(0,it.tex)
            elif k=='_price': e.insert(0,str(it.price))
            elif k=='_weight': e.insert(0,it.weight)
            elif k=='_abun': e.insert(0,str(it.abundance))
            elif k=='_pfx': e.insert(0,it.prefix)
            elif k=='_mcnt': e.insert(0,str(it.mcnt))
            elif k=='_ha': e.insert(0,str(it.ha))
            elif k=='_ba': e.insert(0,str(it.ba))
            elif k=='_la': e.insert(0,str(it.la))
            elif k=='_diff': e.insert(0,str(it.diff))
            elif k=='_hp': e.insert(0,str(it.hp))
            elif k=='_spd': e.insert(0,str(it.spd))
            elif k=='_ms': e.insert(0,str(it.ms))
            elif k=='_wlen': e.insert(0,str(it.wlen))
            elif k=='_ammo': e.insert(0,str(it.ammo))
            elif k=='_thrust': e.insert(0,str(it.thrust))
            elif k=='_swing': e.insert(0,str(it.swing))
            elif k=='_type': e.insert(0,it.type)
            elif k=='_act': e.insert(0,it.act)
            elif k=='_fcnt': e.insert(0,str(it.fcnt))
            elif k=='_tcnt': e.insert(0,str(it.tcnt))
        # Flags
        try: f=int(it.type,16) if it.type.startswith('0x') or len(it.type)>10 else int(it.type)
        except: f=0
        for i,fl in enumerate(ITEM_FLAGS):
            self._flag_vars[fl].set(1 if (f>>(12+i))&1 else 0)
        # Factions
        self._fac_list.delete(0,tk.END)
        for fc in it.facs: self._fac_list.insert(tk.END,str(fc))
        self._status_lbl.config(text=f"Loaded: {it.db}")

    def _apply(self):
        if self._sel<0: return
        try:
            it=self.items[self._sel]
            e=self._entries
            it.db=e['_db'].get(); it.name=e['_name'].get(); it.tex=e['_tex'].get()
            it.price=int(e['_price'].get()or 0); it.weight=e['_weight'].get()
            it.abundance=int(e['_abun'].get()or 0); it.prefix=e['_pfx'].get()
            it.mcnt=int(e['_mcnt'].get()or 0)
            it.ha=int(e['_ha'].get()or 0); it.ba=int(e['_ba'].get()or 0)
            it.la=int(e['_la'].get()or 0); it.diff=int(e['_diff'].get()or 0)
            it.hp=int(e['_hp'].get()or 0); it.spd=int(e['_spd'].get()or 0)
            it.ms=int(e['_ms'].get()or 0); it.wlen=int(e['_wlen'].get()or 0)
            it.ammo=int(e['_ammo'].get()or 0); it.thrust=int(e['_thrust'].get()or 0)
            it.swing=int(e['_swing'].get()or 0)
            it.type=e['_type'].get(); it.act=e['_act'].get()
            it.fcnt=int(e['_fcnt'].get()or 0); it.tcnt=int(e['_tcnt'].get()or 0)
            it.edited=True
            self._fill_list(); self._status_lbl.config(text="Applied ✓")
        except Exception as ex:
            self._status_lbl.config(text=f"Error: {ex}")

    def _reset(self):
        if self._sel>=0: self._on_select(None)

    def _output(self):
        if self._sel<0: return
        it=self.items[self._sel]
        txt=f" {it.db} {it.name} {it.tex} {it.mcnt}"
        for i in range(it.mcnt): txt+=f" {it.mesh_n[i]} {it.mesh_p[i]}" if i<len(it.mesh_n) else " mesh 0"
        txt+=f" {it.type} {it.act} {it.price} {it.prefix} {it.weight} {it.abundance}"
        txt+=f" {it.ha} {it.ba} {it.la} {it.diff} {it.hp} {it.spd} {it.ms} {it.wlen} {it.ammo} {it.thrust} {it.swing}"
        txt+=f" {it.fcnt}"
        for f in it.facs: txt+=f" {f}"
        txt+=f" {it.tcnt}"
        messagebox.showinfo("Output Item",txt)

# ====== TROOP EDITOR ======
class TroopEditor(tk.Frame):
    def __init__(self, parent, repo):
        super().__init__(parent, bg='white'); self.repo=repo; self.troops=repo.troops
        self.pack(fill=tk.BOTH,expand=True); self._sel=-1; self._build()

    def _build(self):
        left=tk.Frame(self,width=360,bg='white')
        left.pack(side=tk.LEFT,fill=tk.Y); left.pack_propagate(False)
        sf=tk.Frame(left,bg='white'); sf.pack(fill=tk.X,padx=4,pady=2)
        self._sv=tk.StringVar()
        tk.Entry(sf,textvariable=self._sv,width=26).pack(side=tk.LEFT)
        tk.Button(sf,text="Find",command=self._filter,width=5).pack(side=tk.LEFT,padx=2)
        self._lb=tk.Listbox(left,width=40,font=('Consolas',9),bg='white')
        self._lb.pack(fill=tk.BOTH,expand=True,padx=4,pady=2); self._lb.bind('<<ListboxSelect>>',self._on_sel)
        self._fill()

        nb=ttk.Notebook(self); nb.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        self._entries={}
        # Basic tab
        bf=tk.Frame(nb); nb.add(bf,text="Basic")
        fields=[("StrID:","_sid"),("Name:","_name"),("Plural:","_plural"),
            ("Flags:","_flags"),("Scene:","_scene"),("Faction:","_faction"),
            ("Upgrade1:","_up1"),("Upgrade2:","_up2"),("Reserved:","_reserved")]
        for i,(l,k) in enumerate(fields):
            tk.Label(bf,text=l,anchor=tk.E,width=12).grid(row=i,column=0,sticky=tk.E,padx=4,pady=1)
            e=tk.Entry(bf,width=28); e.grid(row=i,column=1,sticky=tk.W,padx=4); self._entries[k]=e
        self._up1_lbl=tk.Label(bf,text="",fg='gray'); self._up1_lbl.grid(row=6,column=2,sticky=tk.W)
        self._up2_lbl=tk.Label(bf,text="",fg='gray'); self._up2_lbl.grid(row=7,column=2,sticky=tk.W)

        # Attrs tab
        af=tk.Frame(nb); nb.add(af,text="Attrs/Profs")
        afields=[("Level:","_lv"),("STR:","_str"),("AGI:","_agi"),("INT:","_int"),("CHA:","_cha"),
            ("1H:","_1h"),("2H:","_2h"),("Pole:","_pole"),("Arch:","_arch"),
            ("Xbow:","_xbow"),("Thr:","_thr"),("Fire:","_fire")]
        for i,(l,k) in enumerate(afields):
            tk.Label(af,text=l,anchor=tk.E,width=12).grid(row=i,column=0,sticky=tk.E,padx=4,pady=1)
            e=tk.Entry(af,width=28); e.grid(row=i,column=1,sticky=tk.W,padx=4); self._entries[k]=e

        # Inventory tab
        invf=tk.Frame(nb); nb.add(invf,text="Inventory")
        self._inv_lb=tk.Listbox(invf,font=('Consolas',9),width=50)
        self._inv_lb.pack(fill=tk.BOTH,expand=True,padx=4,pady=4)

        # Skills tab
        skf=tk.Frame(nb); nb.add(skf,text="Skills")
        for i in range(6):
            tk.Label(skf,text=f"Skill {i}:",anchor=tk.E,width=12).grid(row=i,column=0,sticky=tk.E,padx=4,pady=2)
            e=tk.Entry(skf,width=28); e.grid(row=i,column=1,sticky=tk.W,padx=4); self._entries[f'_sk{i}']=e

        # Face tab
        fcf=tk.Frame(nb); nb.add(fcf,text="Face")
        for i in range(8):
            tk.Label(fcf,text=f"Face {i}:",anchor=tk.E,width=12).grid(row=i,column=0,sticky=tk.E,padx=4,pady=2)
            e=tk.Entry(fcf,width=28); e.grid(row=i,column=1,sticky=tk.W,padx=4); self._entries[f'_fc{i}']=e

        # Bottom
        bf2=tk.Frame(self,bg='#f0f0f0',height=36); bf2.pack(side=tk.BOTTOM,fill=tk.X)
        self._sl=tk.Label(bf2,text="",anchor=tk.W,bg='#f0f0f0')
        self._sl.pack(side=tk.LEFT,padx=8,fill=tk.X,expand=True)
        tk.Button(bf2,text="Reset",command=self._reset,width=8).pack(side=tk.RIGHT,padx=4,pady=3)
        tk.Button(bf2,text="Apply",command=self._apply,width=8,bg='#ffaaaa').pack(side=tk.RIGHT,padx=4,pady=3)

    def _fill(self):
        self._lb.delete(0,tk.END)
        for t in self.troops:
            self._lb.insert(tk.END,f"{t.id:4d}  {t.sid:30s} {t.name:20s} Lv{t.lv:3d}")

    def _filter(self):
        q=self._sv.get().lower()
        self._lb.delete(0,tk.END)
        for t in self.troops:
            if not q or q in t.sid.lower() or q in t.name.lower():
                self._lb.insert(tk.END,f"{t.id:4d}  {t.sid:30s} {t.name:20s} Lv{t.lv:3d}")
        self._sl.config(text=f"Filtered: {self._lb.size()}/{len(self.troops)}")

    def _on_sel(self, ev):
        sel=self._lb.curselection()
        if not sel: return
        idx=int(self._lb.get(sel[0]).split()[0])
        self._sel=idx; t=self.troops[idx]
        for k,e in self._entries.items():
            e.delete(0,tk.END)
            val={'_sid':t.sid,'_name':t.name,'_plural':t.plural,'_flags':t.flags,
                '_scene':str(t.scene),'_faction':str(t.faction),'_up1':str(t.up1),
                '_up2':str(t.up2),'_reserved':str(t.reserved),
                '_lv':str(t.lv),'_str':str(t.str),'_agi':str(t.agi),'_int':str(t.int),
                '_cha':str(t.cha)}.get(k,'')
            if not val: val={'_1h':str(t.profs[0]),'_2h':str(t.profs[1]),
                '_pole':str(t.profs[2]),'_arch':str(t.profs[3]),'_xbow':str(t.profs[4]),
                '_thr':str(t.profs[5]),'_fire':str(t.profs[6])}.get(k,'')
            if not val:
                for i in range(6):
                    if k==f'_sk{i}': val=str(t.skills[i]); break
            if not val:
                for i in range(8):
                    if k==f'_fc{i}': val=t.face[i]; break
            e.insert(0,val)
        # Show upgrade names
        if t.up1<len(self.troops): self._up1_lbl.config(text=f"→ {self.troops[t.up1].name}")
        else: self._up1_lbl.config(text="")
        if t.up2<len(self.troops): self._up2_lbl.config(text=f"→ {self.troops[t.up2].name}")
        else: self._up2_lbl.config(text="")
        # Inventory
        self._inv_lb.delete(0,tk.END)
        for i,(iid,qty) in enumerate(t.inv):
            if iid>0:
                name=self.repo.items[iid].db if iid<len(self.repo.items) else f"#{iid}"
                self._inv_lb.insert(tk.END,f"[{i:2d}] {name:35s} x{qty}")
        self._sl.config(text=f"Loaded: {t.sid} Lv{t.lv}")

    def _apply(self):
        if self._sel<0: return
        try:
            t=self.troops[self._sel]; e=self._entries
            t.sid=e['_sid'].get(); t.name=e['_name'].get(); t.plural=e['_plural'].get()
            t.flags=e['_flags'].get(); t.scene=int(e['_scene'].get()or 0)
            t.faction=int(e['_faction'].get()or 0)
            t.up1=int(e['_up1'].get()or 0); t.up2=int(e['_up2'].get()or 0)
            t.reserved=int(e['_reserved'].get()or 0)
            t.lv=int(e['_lv'].get()or 0); t.str=int(e['_str'].get()or 0)
            t.agi=int(e['_agi'].get()or 0); t.int=int(e['_int'].get()or 0)
            t.cha=int(e['_cha'].get()or 0)
            t.profs=[int(e[k].get()or 0) for k in ['_1h','_2h','_pole','_arch','_xbow','_thr','_fire']]
            t.skills=[int(e[f'_sk{i}'].get()or 0) for i in range(6)]
            t.face=[e[f'_fc{i}'].get() for i in range(8)]
            t.edited=True
            self._fill(); self._sl.config(text="Applied ✓")
        except Exception as ex:
            self._sl.config(text=f"Error: {ex}")

    def _reset(self):
        if self._sel>=0: self._on_sel(None)

# ====== FACTION EDITOR ======
class FactionEditor(tk.Frame):
    def __init__(self, parent, repo):
        super().__init__(parent, bg='white'); self.repo=repo; self.factions=repo.factions
        self.pack(fill=tk.BOTH,expand=True); self._sel=-1; self._build()

    def _build(self):
        left=tk.Frame(self,width=320,bg='white')
        left.pack(side=tk.LEFT,fill=tk.Y); left.pack_propagate(False)
        self._lb=tk.Listbox(left,width=36,font=('Consolas',9),bg='white')
        self._lb.pack(fill=tk.BOTH,expand=True,padx=4,pady=4); self._lb.bind('<<ListboxSelect>>',self._on_sel)
        self._fill()

        rf=tk.Frame(self,bg='white'); rf.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        fields=[("StrID:","_sid"),("Name:","_name"),("Flags:","_flags"),("Color:","_color")]
        self._entries={}
        for i,(l,k) in enumerate(fields):
            tk.Label(rf,text=l,anchor=tk.E,width=10).grid(row=i,column=0,sticky=tk.E,padx=6,pady=4)
            e=tk.Entry(rf,width=30); e.grid(row=i,column=1,sticky=tk.W,padx=6); self._entries[k]=e

        tk.Label(rf,text="Relations:",anchor=tk.NW).grid(row=5,column=0,sticky=tk.NW,padx=6,pady=4)
        self._rel_list=tk.Listbox(rf,height=10,width=40); self._rel_list.grid(row=5,column=1,sticky=tk.W)

        bf=tk.Frame(self,bg='#f0f0f0',height=36); bf.pack(side=tk.BOTTOM,fill=tk.X)
        self._sl=tk.Label(bf,text="",anchor=tk.W,bg='#f0f0f0')
        self._sl.pack(side=tk.LEFT,padx=8,fill=tk.X,expand=True)
        tk.Button(bf,text="Apply",command=self._apply,width=8,bg='#ffaaaa').pack(side=tk.RIGHT,padx=4,pady=3)
        tk.Button(bf,text="Reset",command=self._reset,width=8).pack(side=tk.RIGHT,padx=4,pady=3)

    def _fill(self):
        self._lb.delete(0,tk.END)
        for f in self.factions:
            self._lb.insert(tk.END,f"{f.id:3d}  {f.sid:25s} {f.name}")

    def _on_sel(self, ev):
        sel=self._lb.curselection()
        if not sel: return
        idx=int(self._lb.get(sel[0]).split()[0])
        self._sel=idx; f=self.factions[idx]
        for k,e in self._entries.items():
            e.delete(0,tk.END)
            val={'_sid':f.sid,'_name':f.name,'_flags':str(f.flags),'_color':f.color}.get(k,'')
            e.insert(0,val)
        self._rel_list.delete(0,tk.END)
        for i,r in enumerate(f.rels):
            if i<len(self.factions):
                self._rel_list.insert(tk.END,f"[{i:3d}] {self.factions[i].sid:25s} {r:8.4f}")
        self._sl.config(text=f"Loaded: {f.name}")

    def _apply(self):
        if self._sel<0: return
        try:
            f=self.factions[self._sel]; e=self._entries
            f.sid=e['_sid'].get(); f.name=e['_name'].get()
            f.flags=int(e['_flags'].get()or 0); f.color=e['_color'].get()
            f.edited=True
            self._fill(); self._sl.config(text="Applied ✓")
        except Exception as ex:
            self._sl.config(text=f"Error: {ex}")

    def _reset(self):
        if self._sel>=0: self._on_sel(None)

# ====== Main ======
if __name__=="__main__":
    app=App()
    # Auto-open if path provided
    if len(sys.argv)>1 and os.path.isdir(sys.argv[1]):
        load_mod(app.repo, sys.argv[1])
        app._build_tree()
        app._status.config(text=f"Loaded: {len(app.repo.items)} items, {len(app.repo.troops)} troops")
    app.mainloop()
