# SOMATRIN — Améliorer le Bilan Gasoil avec nouveaux graphiques

## Contexte
Projet : C:\Users\pc\Desktop\somatrin_project2
Le bilan gasoil fonctionne déjà sur http://127.0.0.1:8090/gasoil/bilan/
Il faut ajouter 5 nouveaux graphiques et sections.

---

## ÉTAPE 1 — Lire les fichiers existants

Lis ces fichiers avant de toucher quoi que ce soit :
- templates/gasoil/bilan.html → voir la structure actuelle du template
- gasoil/views.py → voir la vue gasoil_bilan() existante et les modèles utilisés
- gasoil/models.py → noter les noms exacts des champs

---

## ÉTAPE 2 — Ajouter les données dans la vue gasoil_bilan()

Dans gasoil/views.py, dans la fonction gasoil_bilan(), ajoute ces nouveaux calculs
APRÈS les calculs existants et AVANT le return render(...) :

```python
# ── 1) Répartition par catégorie d'engin ──────────────────────────────────
# Adapte 'categorie' selon le nom exact du champ dans models.py
# (peut s'appeler categorie, categorie_engin, type_engin, etc.)
from django.db.models import Sum, Count

cat_qs = sorties_filtrees.values('categorie').annotate(
    total=Sum('quantite')
).order_by('-total')[:8]

categories_labels = [str(c['categorie'] or 'Non défini') for c in cat_qs]
categories_values = [round(float(c['total'] or 0), 1) for c in cat_qs]

# ── 2) Consommation par période (semaine par semaine) ─────────────────────
from django.db.models.functions import TruncWeek
semaines_qs = sorties_filtrees.annotate(
    semaine=TruncWeek('date_sortie')
).values('semaine').annotate(
    total=Sum('quantite')
).order_by('semaine')

semaines_labels = [s['semaine'].strftime('%d/%m') if s['semaine'] else '' for s in semaines_qs]
semaines_values = [round(float(s['total'] or 0), 1) for s in semaines_qs]

# ── 3) Top 10 consommation par bon de sortie ──────────────────────────────
# Adapte 'numero_bon' selon le nom exact du champ dans models.py
bons_qs = sorties_filtrees.values('name').annotate(
    total=Sum('quantite')
).order_by('-total')[:10]

bons_labels = [str(b['name'] or 'Sans n°') for b in bons_qs]
bons_values = [round(float(b['total'] or 0), 1) for b in bons_qs]

# ── 4) Top 10 consommation par matricule ─────────────────────────────────
# Adapte 'matricule' selon le nom exact du champ dans models.py
# (peut s'appeler engin, matricule, vehicle, etc.)
matricules_qs = sorties_filtrees.values('matricule').annotate(
    total=Sum('quantite')
).order_by('-total')[:10]

max_mat = float(matricules_qs[0]['total']) if matricules_qs else 1
matricules_data = [
    {
        'nom': str(m['matricule'] or 'Inconnu'),
        'total': round(float(m['total'] or 0), 1),
        'pct': round(float(m['total'] or 0) / max_mat * 100, 1)
    }
    for m in matricules_qs
]

# ── 5) Équipements actifs (matricules uniques sur la période) ─────────────
equipements_qs = sorties_filtrees.values(
    'matricule', 'categorie'
).annotate(
    nb_sorties=Count('id'),
    total_litres=Sum('quantite')
).order_by('-total_litres')[:20]

equipements_actifs = [
    {
        'matricule': str(e['matricule'] or 'Inconnu'),
        'categorie': str(e['categorie'] or '—'),
        'nb_sorties': e['nb_sorties'],
        'total_litres': round(float(e['total_litres'] or 0), 1),
    }
    for e in equipements_qs
]
nb_equipements_actifs = sorties_filtrees.values('matricule').distinct().count()
```

Puis dans le dictionnaire context du return render, ajoute ces clés :
```python
'categories_labels_json': json.dumps(categories_labels),
'categories_values_json': json.dumps(categories_values),
'semaines_labels_json': json.dumps(semaines_labels),
'semaines_values_json': json.dumps(semaines_values),
'bons_labels_json': json.dumps(bons_labels),
'bons_values_json': json.dumps(bons_values),
'matricules_data': matricules_data,
'equipements_actifs': equipements_actifs,
'nb_equipements_actifs': nb_equipements_actifs,
```

IMPORTANT : si les noms de champs ne correspondent pas, lis models.py et adapte.

---

## ÉTAPE 3 — Ajouter les sections dans le template bilan.html

Dans templates/gasoil/bilan.html, ajoute ces sections HTML AVANT la balise {% endblock %} finale.

```html
<!-- ══ LIGNE 3 : Catégorie + Période ══════════════════════════════════ -->
<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;">

  <!-- Répartition par catégorie -->
  <div style="background:#fff;border:0.5px solid #dee2e6;border-radius:14px;padding:22px 24px;">
    <div style="margin-bottom:16px;">
      <div style="font-size:13px;font-weight:700;color:#1a1a2e;">Répartition par catégorie d'engin</div>
      <div style="font-size:11px;color:#8a8f99;margin-top:2px;">Litres consommés par type</div>
    </div>
    <canvas id="chartCategories" height="200"></canvas>
  </div>

  <!-- Consommation par période (semaines) -->
  <div style="background:#fff;border:0.5px solid #dee2e6;border-radius:14px;padding:22px 24px;">
    <div style="margin-bottom:16px;">
      <div style="font-size:13px;font-weight:700;color:#1a1a2e;">Consommation par période</div>
      <div style="font-size:11px;color:#8a8f99;margin-top:2px;">Évolution semaine par semaine</div>
    </div>
    <canvas id="chartPeriode" height="200"></canvas>
  </div>

</div>

<!-- ══ LIGNE 4 : Top matricules ══════════════════════════════════════ -->
<div style="background:#fff;border:0.5px solid #dee2e6;border-radius:14px;padding:22px 24px;margin-bottom:20px;">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;">
    <div>
      <div style="font-size:13px;font-weight:700;color:#1a1a2e;">Top 10 — Consommation par matricule</div>
      <div style="font-size:11px;color:#8a8f99;margin-top:2px;">Litres consommés sur la période sélectionnée</div>
    </div>
  </div>
  <ul style="padding:0;list-style:none;margin:0;">
    {% for m in matricules_data %}
    <li style="display:flex;align-items:center;gap:12px;padding:9px 0;border-bottom:0.5px solid #f5f5f5;">
      <div style="width:22px;height:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0;
        {% if forloop.counter == 1 %}background:#FEF9E7;color:#F39C12;
        {% elif forloop.counter == 2 %}background:#F8F9FA;color:#6c757d;
        {% elif forloop.counter == 3 %}background:#FEF0EA;color:#E8540A;
        {% else %}background:#F8F9FA;color:#8a8f99;{% endif %}">
        {{ forloop.counter }}
      </div>
      <div style="font-size:12px;font-weight:600;color:#1a1a2e;min-width:140px;">{{ m.nom }}</div>
      <div style="flex:1;background:#F5F5F5;border-radius:4px;height:8px;overflow:hidden;">
        <div style="height:100%;border-radius:4px;background:linear-gradient(90deg,#E8540A,#F39C12);width:{{ m.pct }}%;"></div>
      </div>
      <div style="font-size:12px;font-weight:700;color:#1a1a2e;min-width:80px;text-align:right;">{{ m.total }} L</div>
    </li>
    {% empty %}
    <li style="color:#8a8f99;font-size:13px;padding:12px 0;">Aucune donnée disponible</li>
    {% endfor %}
  </ul>
</div>

<!-- ══ LIGNE 5 : Top bons + Équipements actifs ═══════════════════════ -->
<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;">

  <!-- Top bons de sortie -->
  <div style="background:#fff;border:0.5px solid #dee2e6;border-radius:14px;padding:22px 24px;">
    <div style="margin-bottom:16px;">
      <div style="font-size:13px;font-weight:700;color:#1a1a2e;">Top 10 — Consommation par bon</div>
      <div style="font-size:11px;color:#8a8f99;margin-top:2px;">Bons de sortie les plus importants</div>
    </div>
    <canvas id="chartBons" height="250"></canvas>
  </div>

  <!-- Équipements actifs -->
  <div style="background:#fff;border:0.5px solid #dee2e6;border-radius:14px;padding:22px 24px;">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;">
      <div>
        <div style="font-size:13px;font-weight:700;color:#1a1a2e;">Équipements actifs</div>
        <div style="font-size:11px;color:#8a8f99;margin-top:2px;">{{ nb_equipements_actifs }} matricules distincts sur la période</div>
      </div>
      <span style="background:#EAFAF1;color:#27AE60;font-size:11px;font-weight:700;padding:4px 10px;border-radius:10px;">{{ nb_equipements_actifs }} actifs</span>
    </div>
    <div style="overflow-y:auto;max-height:320px;">
      <table style="width:100%;font-size:12px;border-collapse:collapse;">
        <thead>
          <tr style="background:#F8F9FA;">
            <th style="padding:8px 10px;text-align:left;color:#8a8f99;font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:1px;border-bottom:0.5px solid #dee2e6;">Matricule</th>
            <th style="padding:8px 10px;text-align:left;color:#8a8f99;font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:1px;border-bottom:0.5px solid #dee2e6;">Catégorie</th>
            <th style="padding:8px 10px;text-align:center;color:#8a8f99;font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:1px;border-bottom:0.5px solid #dee2e6;">Sorties</th>
            <th style="padding:8px 10px;text-align:right;color:#8a8f99;font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:1px;border-bottom:0.5px solid #dee2e6;">Total (L)</th>
          </tr>
        </thead>
        <tbody>
          {% for e in equipements_actifs %}
          <tr style="border-bottom:0.5px solid #f5f5f5;">
            <td style="padding:8px 10px;font-weight:600;color:#1a1a2e;">{{ e.matricule }}</td>
            <td style="padding:8px 10px;color:#8a8f99;">
              <span style="background:#F8F9FA;border-radius:6px;padding:2px 7px;font-size:10px;">{{ e.categorie }}</span>
            </td>
            <td style="padding:8px 10px;text-align:center;color:#0057A8;font-weight:600;">{{ e.nb_sorties }}</td>
            <td style="padding:8px 10px;text-align:right;font-weight:700;color:#E8540A;">{{ e.total_litres }} L</td>
          </tr>
          {% empty %}
          <tr><td colspan="4" style="padding:16px;text-align:center;color:#8a8f99;">Aucun équipement trouvé</td></tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>

</div>
```

---

## ÉTAPE 4 — Ajouter les scripts Chart.js pour les nouveaux graphiques

Dans le bloc {% block extra_js %} du template bilan.html, après les scripts Chart.js existants, ajoute :

```javascript
// Données nouvelles sections
const catLabels = {{ categories_labels_json|default:'[]'|safe }};
const catValues = {{ categories_values_json|default:'[]'|safe }};
const perLabels = {{ semaines_labels_json|default:'[]'|safe }};
const perValues = {{ semaines_values_json|default:'[]'|safe }};
const bonLabels = {{ bons_labels_json|default:'[]'|safe }};
const bonValues = {{ bons_values_json|default:'[]'|safe }};

const pal10 = ['#E8540A','#0057A8','#27AE60','#F39C12','#8E44AD','#16A085','#E74C3C','#2980B9','#D35400','#1ABC9C'];
const tipStyle = {backgroundColor:'#fff',titleColor:'#1a1a2e',bodyColor:'#8a8f99',borderColor:'#dee2e6',borderWidth:1,padding:10};

// Graphique catégories (doughnut)
if (document.getElementById('chartCategories')) {
  new Chart(document.getElementById('chartCategories'), {
    type: 'doughnut',
    data: {
      labels: catLabels.length ? catLabels : ['Camions','Pelles','Chargeuses','Groupes','Autres'],
      datasets: [{
        data: catValues.length ? catValues : [8200,4500,3100,1800,900],
        backgroundColor: pal10,
        borderWidth: 3,
        borderColor: '#fff',
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true,
      cutout: '60%',
      plugins: {
        legend: { position: 'right', labels: { usePointStyle: true, pointStyle: 'circle', padding: 10, font: { size: 11 } } },
        tooltip: { ...tipStyle, callbacks: { label: c => ` ${c.label}: ${c.parsed.toLocaleString('fr-FR')} L` } }
      }
    }
  });
}

// Graphique période (area chart)
if (document.getElementById('chartPeriode')) {
  new Chart(document.getElementById('chartPeriode'), {
    type: 'line',
    data: {
      labels: perLabels.length ? perLabels : ['S1','S2','S3','S4','S5','S6','S7','S8'],
      datasets: [{
        label: 'Litres',
        data: perValues.length ? perValues : [12000,18000,9000,22000,15000,19000,11000,16000],
        borderColor: '#0057A8',
        backgroundColor: 'rgba(0,87,168,0.08)',
        borderWidth: 2.5,
        pointRadius: 4,
        pointBackgroundColor: '#0057A8',
        fill: true,
        tension: 0.4
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { ...tipStyle, callbacks: { label: c => ` ${c.parsed.y.toLocaleString('fr-FR')} L` } }
      },
      scales: {
        x: { grid: { display: false }, border: { display: false } },
        y: { grid: { color: '#F5F5F5' }, border: { display: false }, ticks: { callback: v => v.toLocaleString('fr-FR') } }
      }
    }
  });
}

// Graphique bons (bar horizontal)
if (document.getElementById('chartBons')) {
  new Chart(document.getElementById('chartBons'), {
    type: 'bar',
    data: {
      labels: bonLabels.length ? bonLabels : ['BON001','BON002','BON003','BON004','BON005'],
      datasets: [{
        label: 'Litres',
        data: bonValues.length ? bonValues : [4500,3800,3200,2900,2500],
        backgroundColor: pal10,
        borderRadius: 5,
        borderSkipped: false
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { ...tipStyle, callbacks: { label: c => ` ${c.parsed.x.toLocaleString('fr-FR')} L` } }
      },
      scales: {
        x: { grid: { color: '#F5F5F5' }, border: { display: false }, ticks: { callback: v => v.toLocaleString('fr-FR') } },
        y: { grid: { display: false }, border: { display: false }, ticks: { font: { size: 10 } } }
      }
    }
  });
}
```

---

## ÉTAPE 5 — Vérification des noms de champs

AVANT d'écrire le code final, vérifie dans models.py :
- Le champ "catégorie" s'appelle comment exactement ?
- Le champ "matricule/engin" s'appelle comment exactement ?
- Le champ "numéro de bon" s'appelle comment exactement ?
- Le champ "date" s'appelle comment exactement ? (date, date_sortie, date_mouvement ?)

Adapte TOUS les noms de champs dans le code de l'étape 2 selon les vrais noms.

---

## ÉTAPE 6 — Test final

Lance le serveur :
python manage.py runserver 8090 --settings=somatrin.settings_local

Vérifie que http://127.0.0.1:8090/gasoil/bilan/ s'affiche sans erreur avec :
- Le graphique catégories (donut)
- Le graphique période (courbe semaines)
- Le top 10 matricules (barres horizontales avec dégradé)
- Le graphique bons (barres horizontales)
- Le tableau équipements actifs

Si erreur, montre le traceback et corrige.
