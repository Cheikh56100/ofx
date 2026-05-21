# 🏦 BanqScan — Convertisseur OCR de relevés bancaires

Application Streamlit pour convertir des relevés bancaires scannés en PDF texte + CSV.

## ✨ Deux modes

| | Mode Gratuit | Mode Premium |
|---|---|---|
| **OCR** | Tesseract (local) | Claude AI (Anthropic) |
| **Coût** | 0 € | ~0.01 € / page |
| **Clé API** | ❌ Non requise | ✅ Requise |
| **Précision** | Bonne | Excellente |
| **Formats** | Tous | Tous |

## 🚀 Déploiement Streamlit Cloud (gratuit)

### Étape 1 — Fichiers nécessaires

Déposez ces fichiers dans un repo GitHub :
```
banqscan/
├── app.py
├── requirements.txt
└── packages.txt        ← IMPORTANT pour Tesseract
```

### Étape 2 — Déployer

1. [share.streamlit.io](https://share.streamlit.io) → **New app**
2. Sélectionnez votre repo → fichier `app.py`
3. **Deploy** ✅

### Étape 3 — Clé API (optionnel, pour Mode Premium)

Dans Streamlit Cloud → **Settings → Secrets** :
```toml
ANTHROPIC_API_KEY = "sk-ant-votre-cle"
```

## 💻 Lancement local

```bash
# Installer Tesseract (Ubuntu/Debian)
sudo apt-get install tesseract-ocr tesseract-ocr-fra

# Installer Python deps
pip install -r requirements.txt

# Lancer
streamlit run app.py
```

## 📁 Formats d'image acceptés

JPEG · JFIF · JPG · PNG · BMP · TIFF · WEBP · GIF · TGA · AVIF · PPM · PGM…

## 📦 Exports

- **PDF** : texte 100% sélectionnable, lisible par votre logiciel
- **CSV** : séparateur `;`, BOM UTF-8, compatible Excel français, Sage, EBP, Cegid…
