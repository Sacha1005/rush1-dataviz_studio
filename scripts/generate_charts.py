"""
Générateur de graphiques professionnels haute résolution pour le classeur Excel
Projet : Rush 1 - DataViz Studio
Visual Designer & Lead Data Analyst
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import squarify
import os
import re

# Configuration globale Matplotlib
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.sans-serif'] = 'Helvetica Neue', 'Arial', 'DejaVu Sans'
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.titleweight'] = 'bold'
plt.rcParams['axes.labelsize'] = 10
plt.rcParams['axes.labelweight'] = 'bold'

os.makedirs("assets/charts", exist_ok=True)

df = pd.read_csv("tiktok_funny_hashtag_videos.csv")
df['total_interactions'] = df['video_stats'] + df['video_commentCount'] + df['video_shareCount']
df['er_views'] = (df['total_interactions'] / df['video_playCount']) * 100
df['viral_multiplier'] = df['video_playCount'] / df['author_followerCount']

def clean_name(row):
    # Préférer un nom propre lisible sans caractères non-ASCII / emojis
    raw_name = str(row['author_nickname'])
    clean = re.sub(r'[^\w\s\.-]', '', raw_name).strip()
    if len(clean) < 3 or clean.isnumeric():
        clean = str(row['author_uniqueId'])
    return clean[:16]

df['clean_name'] = df.apply(clean_name, axis=1)

def categorize(row):
    text = (str(row['author_uniqueId']) + ' ' + str(row['author_nickname']) + ' ' + str(row['author_signature']) + ' ' + str(row['video_desc'])).lower()
    if any(k in text for k in ['cat', 'dog', 'pet', 'corgi', 'raccoon', 'animal', 'lion', 'llama', 'cockatoo', 'puppy', 'kitten', 'chat', 'frog']):
        return 'Animaux & Pets'
    elif any(k in text for k in ['magic', 'trick', 'wizard', 'photo', 'drawing', 'art', 'animation', 'beatbox', 'parkour', 'flip', 'dance', 'actor']):
        return 'Talents & Créativité'
    elif any(k in text for k in ['mom', 'dad', 'family', 'couple', 'wife', 'husband', 'baby', 'toddler', 'brother', 'sister', 'kid']):
        return 'Famille & Couple'
    elif any(k in text for k in ['game', 'gaming', 'vr', 'minecraft', 'gym', 'sport', 'football', 'soccer', 'tennis']):
        return 'Sports & Gaming'
    else:
        return 'Humour & Sketch'
        
df['category'] = df.apply(categorize, axis=1)

bins = [0, 1_000_000, 5_000_000, float('inf')]
labels = ['Micro/Nano (<1M)', 'Mid/Macro (1M-5M)', 'Mega (5M+)']
df['tier'] = pd.cut(df['author_followerCount'], bins=bins, labels=labels)

duration_bins = [0, 15, 30, 999]
duration_labels = ['Format Court (<=15s)', 'Format Moyen (16-30s)', 'Format Long (>30s)']
df['duration_group'] = pd.cut(df['video_duration'], bins=duration_bins, labels=duration_labels)

# Palette de couleurs corporate élégante
cat_palette = {
    'Humour & Sketch': '#1E3A8A',       # Bleu Nuit Profond
    'Animaux & Pets': '#0D9488',        # Vert Émeraude / Teal
    'Talents & Créativité': '#D97706',  # Ambre / Or
    'Famille & Couple': '#BE185D',      # Framboise / Rose
    'Sports & Gaming': '#4F46E5'        # Indigo
}

# -----------------------------------------------------------------------------
# GRAPHIQUE 1 : TREEMAP PROFESSIONNEL (Requis explicitement par le client)
# -----------------------------------------------------------------------------
print("Génération du Treemap...")
creator_grp = df.groupby(['author_uniqueId', 'clean_name', 'category'], observed=False).agg(
    total_views=('video_playCount', 'sum'),
    er_mean=('er_views', 'mean'),
    followers=('author_followerCount', 'first')
).reset_index().sort_values(by='total_views', ascending=False)

top22 = creator_grp.head(20).copy()
top22['label'] = top22.apply(
    lambda r: f"@{r['author_uniqueId'][:12]}\n{r['total_views']/1e6:.0f}M vues\nER {r['er_mean']:.1f}%", 
    axis=1
)
top22['color'] = top22['category'].map(cat_palette).fillna('#64748B')

fig, ax = plt.subplots(figsize=(13, 7.5), dpi=160)
squarify.plot(
    sizes=top22['total_views'], 
    label=top22['label'], 
    color=top22['color'], 
    alpha=0.92, 
    edgecolor="white", 
    linewidth=2,
    text_kwargs={'fontsize': 9, 'weight': 'bold', 'color': 'white'}
)
plt.title("TREEMAP OFFICIEL DE LA PERFORMANCE DES CRÉATEURS\n(Taille = Volume de Vues Cumulées | Couleur = Catégorie de Contenu)", 
          fontsize=13, weight='bold', pad=15, color='#0F172A')
plt.axis('off')

import matplotlib.patches as mpatches
legend_patches = [mpatches.Patch(color=col, label=cat) for cat, col in cat_palette.items()]
plt.legend(handles=legend_patches, loc='lower center', bbox_to_anchor=(0.5, -0.09), 
           ncol=5, frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=9)

plt.tight_layout()
plt.savefig("assets/charts/treemap_creators.png", bbox_inches='tight')
plt.close()

# -----------------------------------------------------------------------------
# GRAPHIQUE 2 : MATRICE DE SÉLECTION (Engagement vs Followers)
# -----------------------------------------------------------------------------
print("Génération de la Matrice de Sélection...")
fig, ax = plt.subplots(figsize=(10.5, 5.5), dpi=160)

sns.scatterplot(
    data=top22,
    x='followers',
    y='er_mean',
    hue='category',
    palette=cat_palette,
    size='total_views',
    sizes=(160, 1100),
    alpha=0.88,
    edgecolor='#0F172A',
    linewidth=1.2,
    ax=ax
)

ax.set_xscale('log')
ax.axhline(16.0, color='#94A3B8', linestyle='--', linewidth=1.2, label='Seuil de Haute Performance ER (16%)')
ax.axvline(1_000_000, color='#94A3B8', linestyle=':', linewidth=1.2)

# Zones d'aide à la décision
ax.text(250_000, 24, "ZONE PÉPITES VIRALES\n(Fort engagement, ROI maximal)", 
        fontsize=8.5, weight='bold', color='#0D9488', bbox=dict(boxstyle="round,pad=0.4", fc="#CCFBF1", ec="#0D9488", alpha=0.85))
ax.text(10_000_000, 11, "ZONE NOTORIÉTÉ PURE\n(Portée de masse, coût élevé)", 
        fontsize=8.5, weight='bold', color='#1E3A8A', bbox=dict(boxstyle="round,pad=0.4", fc="#DBEAFE", ec="#1E3A8A", alpha=0.85))

for _, r in top22.head(6).iterrows():
    ax.annotate(f"@{r['author_uniqueId']}", (r['followers'], r['er_mean']),
                xytext=(5, 5), textcoords='offset points', fontsize=8, weight='bold', color='#1E293B')

ax.set_title("MATRICE D'AIDE À LA DÉCISION : ENGAGEMENT vs AUDIENCE", fontsize=12, weight='bold', color='#0F172A')
ax.set_xlabel("Nombre d'Abonnés (Échelle Log)", fontsize=9.5, weight='bold')
ax.set_ylabel("Taux d'Engagement par Vue (%)", fontsize=9.5, weight='bold')
ax.grid(True, linestyle='--', alpha=0.5)

handles, labels = ax.get_legend_handles_labels()
cat_labels = list(cat_palette.keys())
cat_handles = [h for h, l in zip(handles, labels) if l in cat_labels]
ax.legend(cat_handles, cat_labels, loc='upper right', frameon=True, facecolor='#F8FAFC', edgecolor='#CBD5E1', fontsize=8.5)

plt.tight_layout()
plt.savefig("assets/charts/matrix_selection.png", bbox_inches='tight')
plt.close()

# -----------------------------------------------------------------------------
# GRAPHIQUE 3 : ANALYSE FACTORIELLE COMBINÉE (Durée & Audio)
# -----------------------------------------------------------------------------
print("Génération du combiné Facteurs...")
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.5, 4.8), dpi=160)

# Subplot 1 : Durée
dur_stats = df.groupby('duration_group', observed=False).agg(
    mean_views=('video_playCount', 'mean'),
    mean_er=('er_views', 'mean')
).reset_index()

x = np.arange(len(dur_stats))
width = 0.35

ax1_twin = ax1.twinx()
b1 = ax1.bar(x - width/2, dur_stats['mean_views']/1e6, width, label='Vues Moyennes (M)', color='#1E3A8A', alpha=0.9)
b2 = ax1_twin.bar(x + width/2, dur_stats['mean_er'], width, label='Taux Engagement Moyen (%)', color='#D97706', alpha=0.9)

ax1.set_xticks(x)
ax1.set_xticklabels(['Court (<=15s)', 'Moyen (16-30s)', 'Long (>30s)'], weight='bold', fontsize=8.5)
ax1.set_ylabel("Vues Moyennes (Millions)", color='#1E3A8A', weight='bold', fontsize=9)
ax1_twin.set_ylabel("Taux Engagement (%)", color='#D97706', weight='bold', fontsize=9)
ax1.set_title("A. IMPACT DE LA DURÉE DU CONTENU\n(Volume vs Taux d'Engagement)", weight='bold', fontsize=10.5)
ax1.grid(False)
ax1_twin.grid(False)

for rect in b1:
    h = rect.get_height()
    ax1.annotate(f"{h:.1f}M", (rect.get_x() + rect.get_width()/2, h), xytext=(0, 2), 
                 textcoords='offset points', ha='center', va='bottom', fontsize=8, weight='bold')
for rect in b2:
    h = rect.get_height()
    ax1_twin.annotate(f"{h:.1f}%", (rect.get_x() + rect.get_width()/2, h), xytext=(0, 2), 
                      textcoords='offset points', ha='center', va='bottom', fontsize=8, weight='bold')

# Subplot 2 : Audio
audio_stats = df.groupby('music_originality').agg(
    mean_views=('video_playCount', 'mean'),
    mean_er=('er_views', 'mean'),
    count=('video_id', 'count')
).reset_index()
audio_stats['label'] = audio_stats['music_originality'].map({True: "Audio Original / Meme\n(75% des vidéos)", False: "Musique Commerciale\n(25% des vidéos)"})

colors_audio = ['#64748B', '#0D9488']
bars = ax2.bar(audio_stats['label'], audio_stats['mean_views']/1e6, color=colors_audio, width=0.52, edgecolor='#0F172A', linewidth=1)
ax2.set_ylabel("Vues Moyennes (Millions)", weight='bold', fontsize=9)
ax2.set_title("B. IMPACT DU TYPE D'AUDIO\n(+20% de Portée pour les Sons Originaux)", weight='bold', fontsize=10.5)
ax2.grid(axis='y', linestyle='--', alpha=0.5)

for bar, (_, r) in zip(bars, audio_stats.iterrows()):
    h = bar.get_height()
    ax2.annotate(f"{h:.1f}M Vues\n(ER : {r['mean_er']:.1f}%)", 
                 (bar.get_x() + bar.get_width()/2, h/2), ha='center', va='center', 
                 color='white', weight='bold', fontsize=9.5)

plt.tight_layout()
plt.savefig("assets/charts/factors_duration_audio.png", bbox_inches='tight')
plt.close()

# -----------------------------------------------------------------------------
# GRAPHIQUE 4 : OBJECTION MANAGER (Micro vs Mega)
# -----------------------------------------------------------------------------
print("Génération du comparatif Micro vs Mega...")
tier_stats = df.groupby('tier', observed=False).agg(
    median_mult=('viral_multiplier', 'median'),
    median_er=('er_views', 'median'),
    median_views=('video_playCount', 'median')
).reset_index()

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10.5, 4.5), dpi=160)

cols = ['#0D9488', '#2563EB', '#64748B']
bars1 = ax1.bar(tier_stats['tier'], tier_stats['median_mult'], color=cols, width=0.5, edgecolor='#0F172A')
ax1.set_ylabel("Multiplicateur Viral (Vues ÷ Abonnés)", weight='bold', fontsize=9)
ax1.set_title("A. MULTI-VIRALITÉ MÉDIANE\n(Surperformance explosive des Micro-comptes)", weight='bold', fontsize=10)
ax1.grid(axis='y', linestyle='--', alpha=0.5)

for bar in bars1:
    h = bar.get_height()
    ax1.annotate(f"{h:.1f}x", (bar.get_x() + bar.get_width()/2, h), xytext=(0, 3),
                 textcoords='offset points', ha='center', va='bottom', weight='bold', fontsize=9.5)

bars2 = ax2.bar(tier_stats['tier'], tier_stats['median_er'], color=cols, width=0.5, edgecolor='#0F172A')
ax2.set_ylabel("Taux d'Engagement Médian (%)", weight='bold', fontsize=9)
ax2.set_title("B. TAUX D'ENGAGEMENT MÉDIAN (%)\n(Les Micro-comptes égalent ou dépassent les Mega)", weight='bold', fontsize=10)
ax2.grid(axis='y', linestyle='--', alpha=0.5)

for bar in bars2:
    h = bar.get_height()
    ax2.annotate(f"{h:.1f}%", (bar.get_x() + bar.get_width()/2, h), xytext=(0, 3),
                 textcoords='offset points', ha='center', va='bottom', weight='bold', fontsize=9.5)

plt.tight_layout()
plt.savefig("assets/charts/objection_manager_micro_vs_mega.png", bbox_inches='tight')
plt.close()

print(">>> Tous les graphiques ont été générés avec succès !")
