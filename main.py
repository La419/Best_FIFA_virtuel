import os
import re
import math
import json
import base64
import time
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from telegram import Update
from telegram.ext import ApplicationBuilder, ContextTypes, MessageHandler, filters

# ==========================================
# 1. CONFIGURATION (Config)
# ==========================================
class Config:
    TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "9144033144:AAEhbbpii7PYhW94Lo_l0atLU1CHIJJlE4Y")
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "AQ.Ab8RN6IEP0kpIqK2LpEZPw3BDJco56z8F5hC7fV9_evzI01ElQ")
    
    GEMINI_MODEL = "gemini-3.8-flash"
    SEUIL_CONFIANCE_MIN = 80.0
    
    SYSTEM_PROMPT = (
        "Tu es un expert en analyse statistique de matchs virtuels FIFA et de paris sportifs. "
        "Analyse cette image de matchs virtuels et extrait les données sous la forme stricte d'un tableau JSON d'objets avec les clés :\n"
        "- 'championnat' (string)\n"
        "- 'equipe_domicile' (string)\n"
        "- 'equipe_exterieur' (string)\n"
        "- 'buts_dom_historique' (liste d'entiers des buts récents de l'équipe domicile)\n"
        "- 'buts_ext_historique' (liste d'entiers des buts récents de l'équipe extérieur)\n"
        "- 'cotes_bookmaker' (un objet contenant par exemple: {'v1': float, 'nul': float, 'v2': float})\n"
        "- 'est_rejoue' (booléen : True si la confrontation est rejouée, False sinon)\n"
        "Règle spécifique pour les matchs rejoués : Si une confrontation est rejouée, pars de l'hypothèse qu'au moins une des deux mi-temps dépasse 6.5 buts (7 ou plus).\n"
        "Renvoie UNIQUEMENT le JSON brut sans balises Markdown."
    )

# ==========================================
# 2. LOGIQUE D'ANALYSE MATHÉMATIQUE
# ==========================================
class AnalyseurVirtuel:
    def _init_(self, seuil_confiance_min=Config.SEUIL_CONFIANCE_MIN):
        self.seuil_confiance_min = seuil_confiance_min

    @staticmethod
    def poisson_pmf(k, lam):
        if lam <= 0:
            return 0.0
        return (math.exp(-lam) * (lam ** k)) / math.factorial(k)

    @staticmethod
    def poisson_cdf(k_max, lam):
        return sum(AnalyseurVirtuel.poisson_pmf(i, lam) for i in range(k_max + 1))

    def calculer_lambda(self, historique):
        if not historique:
            return 1.2
        return sum(historique) / len(historique)

    def evaluer_domination_et_poisson(self, hist_dom, hist_ext):
        lam_dom = self.calculer_lambda(hist_dom)
        lam_ext = self.calculer_lambda(hist_ext)

        p_dom_2plus = 1 - self.poisson_cdf(1, lam_dom)
        p_ext_2plus = 1 - self.poisson_cdf(1, lam_ext)

        total_lam = lam_dom + lam_ext if (lam_dom + lam_ext) > 0 else 1.0
        confiance_v1 = round((lam_dom / total_lam) * 100, 2)
        confiance_v2 = round((lam_ext / total_lam) * 100, 2)

        return {
            "lambda_dom": round(lam_dom, 2),
            "lambda_ext": round(lam_ext, 2),
            "confiance_v1": confiance_v1,
            "confiance_v2": confiance_v2
        }

    def analyser_match(self, match_data):
        championnat = match_data.get("championnat", "Inconnu")
        dom = match_data.get("equipe_domicile", "Équipe A")
        ext = match_data.get("equipe_exterieur", "Équipe B")
        hist_dom = match_data.get("buts_dom_historique", [])
        hist_ext = match_data.get("buts_ext_historique", [])
        cotes_bk = match_data.get("cotes_bookmaker", {"v1": 2.0, "nul": 3.2, "v2": 3.0})
        est_rejoue = match_data.get("est_rejoue", False)

        stats = self.evaluer_domination_et_poisson(hist_dom, hist_ext)
        confiance_max = max(stats["confiance_v1"], stats["confiance_v2"])
        
        vainqueur_mt1 = "V1" if stats["lambda_dom"] > stats["lambda_ext"] else ("V2" if stats["lambda_ext"] > stats["lambda_dom"] else "Nul")
        vainqueur_mt2 = "V2" if stats["lambda_dom"] <= stats["lambda_ext"] else "V1"
        vainqueur_full = "V1" if stats["confiance_v1"] > stats["confiance_v2"] else "V2"

        strategie = "REJOUÉ : Hypothèse MT > 6.5 buts active" if est_rejoue else "Standard"

        if confiance_max > self.seuil_confiance_min:
            if stats["confiance_v1"] > stats["confiance_v2"]:
                decision = f"<b>PARIER V1</b> (Value Bet - {stats['confiance_v1']}%)"
            else:
                decision = f"<b>PARIER V2</b> (Value Bet - {stats['confiance_v2']}%)"
        else:
            decision = f"<s>PASSER</s> (Confiance {confiance_max}% <= {self.seuil_confiance_min}%)"

        return {
            "championnat": championnat,
            "match": f"{dom} vs {ext}",
            "cotes_bookmaker": str(cotes_bk),
            "vainqueur_mt1": vainqueur_mt1,
            "vainqueur_mt2": vainqueur_mt2,
            "vainqueur_full": vainqueur_full,
            "strategie": strategie,
            "confiance": f"{confiance_max}%",
            "decision": decision
        }

analyseur = AnalyseurVirtuel()

# ==========================================
# 3. INTERACTION GEMINI VISION
# ==========================================
def analyser_image_avec_gemini(image_path, max_retries=3):
    with open(image_path, "rb") as f:
        image_bytes = f.read()
    image_base64 = base64.b64encode(image_bytes).decode("utf-8")

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{Config.GEMINI_MODEL}:generateContent?key={Config.GEMINI_API_KEY}"
    
    payload = {
        "contents": [{
            "parts": [
                {"text": Config.SYSTEM_PROMPT},
                {"inline_data": {"mime_type": "image/jpeg", "data": image_base64}}
            ]
        }]
    }

    headers = {"Content-Type": "application/json"}

    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(url, json=payload, headers=headers, timeout=30)
            if response.status_code == 200:
                result_json = response.json()
                return result_json["candidates"][0]["content"]["parts"][0]["text"]
            else:
                raise Exception(f"Erreur API Gemini ({response.status_code}): {response.text}")
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
            if attempt == max_retries:
                raise Exception(f"Échec de connexion après {max_retries} tentatives : {e}")
            time.sleep(2)

# ==========================================
# 4. GESTION TELEGRAM & EXCEL (Format HTML)
# ==========================================
async def handle_image(update: Update, context: ContextTypes.DEFAULT_TYPE):
    photo_file = await update.message.photo[-1].get_file()
    photo_path = "temp_match.jpg"
    await photo_file.download_to_drive(photo_path)
    
    await update.message.reply_text("📊 <b>Analyse en cours...</b> Calcul de Poisson et vérification des >80%.", parse_mode="HTML")

    try:
        texte_reponse = analyser_image_avec_gemini(photo_path)
        texte_reponse = texte_reponse.strip()
        texte_reponse = re.sub(r"^json\s*|\s*$", "", texte_reponse, flags=re.IGNORECASE)
        matchs_data = json.loads(texte_reponse)
    except Exception as e:
        await update.message.reply_text(f"❌ <b>Erreur :</b> {e}", parse_mode="HTML")
        if os.path.exists(photo_path):
            os.remove(photo_path)
        return

    donnees_calculs = []
    for match in matchs_data:
        res = analyseur.analyser_match(match)
        donnees_calculs.append([
            res["championnat"], res["match"], res["cotes_bookmaker"],
            res["vainqueur_mt1"], res["vainqueur_mt2"], res["vainqueur_full"],
            res["strategie"], res["confiance"], res["decision"]
        ])

    if not donnees_calculs:
        await update.message.reply_text("⚠️ Aucun match exploitable trouvé.")
        if os.path.exists(photo_path):
            os.remove(photo_path)
        return

    output_excel = "Analyse_Comparaison_Cotes.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Décisions & Cotes FIFA"

    headers = ["Championnat", "Match", "Côtes Bookmaker", "MT1", "MT2", "Full Time", "Stratégie", "Confiance", "Décision Finale"]
    ws.append(headers)
    for row in donnees_calculs:
        ws.append(row)

    # Styles Excel
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    for row_idx in range(2, len(donnees_calculs) + 2):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = center_align

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_excel)

    with open(output_excel, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename="Analyse_Comparaison_Cotes.xlsx",
            caption="✅ <b>Analyse terminée avec succès !</b>\n📊 Rapport Excel et filtres de confiance appliqués.",
            parse_mode="HTML"
        )

    if os.path.exists(photo_path):
        os.remove(photo_path)
    if os.path.exists(output_excel):
        os.remove(output_excel)

# ==========================================
# 5. LANCEMENT DU BOT
# ==========================================
def main():
    app = ApplicationBuilder().token(Config.TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.PHOTO, handle_image))
    print("🤖 Bot Telegram en cours d'exécution...")
    app.run_polling()

if _name_ == '_main_':
    main()