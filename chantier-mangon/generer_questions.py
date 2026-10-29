"""Génère la liste de toutes les questions ouvertes du chantier Mangon (Word puis PDF)."""
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Cm, Pt

from donnees import CONSTATS, PIECES
from generer_bilan import ICI, PHOTOS, STATUTS, COUL_STATUT, ecrire, ligne, ombrer, tableau

SORTIE = ICI / "Questions_electricite_Mangon_2026-10-29.docx"

ELEC, ISEO = "Électricien", "ISEO"

# id du point -> (destinataire, question). Les points sans question ici ne sont pas repris.
QUESTIONS = {
    "tableau--diff": (ELEC, "Les différentiels posés sont bipolaires alors que le devis prévoit des tétrapolaires, et il n'y a ni 25 A ni 6 A. Pourquoi cet écart ? Une moins-value est-elle prévue ? Quel est le problème de disjoncteur constaté ?"),
    "tableau--compteur": (ELEC, "Le compteur est-il monophasé ou triphasé, et de quelle puissance ?"),
    "tableau--listing": (ELEC, "Pouvez-vous remettre un listing du tableau qui reprend la numérotation des chambres du devis ?"),
    "tableau--terre": (ELEC, "Pouvez-vous fournir le PV de mesure de la terre (9,6 Ω relevés) ?"),
    "tableau--equipot": (ELEC, "Les liaisons équipotentielles principales et secondaires (salles d'eau) sont-elles toutes réalisées ?"),
    "tableau--test": (ELEC, "Chaque différentiel a-t-il été testé avec son bouton test ?"),
    "tableau--niko": (ELEC, "Quelle gamme et quelle couleur NIKO ont été posées ? Le supplément Intense (blanc à l'étage, anthracite au rez) est-il appliqué ?"),
    "halls--spots": (ELEC, "Le devis prévoit 8 spots, 6 sont posés : les 2 spots en moins seront-ils déduits ? Comment le prolongement visible sur les photos a-t-il été réalisé ?"),
    "douche--points": (ELEC, "Le devis prévoit 2 points lumineux au plafond, les plans 1 : combien sont posés et pourquoi cette différence ?"),
    "douche--douche": (ELEC, "La prise double devait être indépendante : pourquoi ne fonctionne-t-elle que lumière allumée ? L'interrupteur est-il du bon côté ? Le point au-dessus de la douche (plan envoyé) est-il prévu ?"),
    "douche--equipot": (ELEC, "La liaison équipotentielle de la salle de douche est-elle réalisée ?"),
    "garde--spots": (ELEC, "3 spots au devis, 2 posés : le 3e sera-t-il posé ou déduit ?"),
    "garde--detect": (ELEC, "Le détecteur posé n'est pas au devis : est-il facturé en supplément ?"),
    "cuisine--spots": (ELEC, "6 spots posés pour 4 au devis : les 2 spots supplémentaires sont-ils facturés ?"),
    "cuisine--inter": (ELEC, "L'interrupteur des spots n'est pas à l'endroit demandé : quand sera-t-il déplacé ?"),
    "cuisine--meuble": (ELEC, "Le point lumineux sous meuble est au devis mais n'est ni sur les plans ni posé : quand sera-t-il réalisé ?"),
    "cuisine--pdt": (ISEO, "Le plan envoyé prévoyait une prise sous la fenêtre en plus des 2 prises doubles : la veut-on ? Si oui, demander le prix."),
    "buanderie--appl": (ELEC, "Les appliques extérieures sont en 2 directions sans témoin, alors que le devis prévoit 1 direction + témoin : le témoin sera-t-il ajouté ?"),
    "buanderie--ext": (ELEC, "La prise extérieure ne correspond pas au devis (NIKO Hydro noire) : sera-t-elle remplacée ?"),
    "salon--appl": (ELEC, "Les 2 appliques murales en 2 directions (devis et plans, A3) ne sont pas posées : quand le seront-elles ?"),
    "sam--inter": (ISEO, "Une erreur a été notée sur les interrupteurs près de la porte du hall (photo) : laquelle exactement ? À préciser avant de la signaler."),
    "garage--tube": (ELEC, "Le tube LED doit être placé au plus haut : pouvez-vous le remonter ?"),
    "garage--prise": (ELEC, "La prise NIKO Hydro est grise au lieu de noire : sera-t-elle remplacée ?"),
    "garage--chaud": (ELEC, "L'alimentation de la chaudière (hors devis, ajoutée aux plans) n'est pas conforme : quel est le problème ? Est-elle facturée ?"),
    "garage--pac": (ELEC, "Les 3 alimentations au devis (pompe à chaleur int. + ext. et chauffe-eau) ne sont pas conformes : quand seront-elles réalisées correctement ?"),
    "garage--td": (ELEC, "Le tableau du garage n'est pas au devis : est-il facturé en supplément, et à quel prix ?"),
    "garage--jardin": (ISEO, "L'éclairage extérieur côté jardin figurait sur le plan envoyé mais pas au devis : le veut-on ? Si oui, demander le prix."),
    "hall12--eclairage": (ELEC, "Le devis prévoit 5 points classiques en 3 directions, les plans 5 spots, et ce qui est posé ne correspond pas à ce qui avait été convenu (5 spots). Qu'est-ce qui est posé exactement, et y a-t-il un surcoût pour les spots ?"),
    "hall12--commandes": (ELEC, "Les 3 commandes de l'escalier (pied, haut, hall) ont-elles bien leur témoin, comme au devis ?"),
    "hall34--rail": (ELEC, "Le point sur rail est posé (F2) mais pas le rail : qui fournit et pose le rail ?"),
    "hall34--appl": (ELEC, "Le devis prévoit une applique en 2 directions + témoin : pourquoi le témoin manque-t-il ? La commande est-elle bien en 2 directions ? Où est l'applique ?"),
    "sdb--equipot": (ELEC, "La liaison équipotentielle de la salle de bains est-elle réalisée ? Pouvez-vous en fournir la preuve (contrôle de l'organisme agréé) ?"),
    "ch1--doubles": (ISEO, "Les appliques et interrupteurs de tête de lit devaient pouvoir éteindre la lampe de la chambre. Au devis, c'est prévu en chambre 3 (16,63 m²), pas en chambre 1 : dans quelle chambre le veut-on ? Si c'est la chambre 1, demander le coût de la modification."),
    "ch2--doubles": (ELEC, "3 prises doubles au devis et aux plans, 2 constatées : où est la 3e ?"),
    "ch3--simples": (ELEC, "Les prises simples de la chambre 3 (3 au devis, 2 aux plans) sont absentes : quand seront-elles posées ?"),
    "ch3--usb": (ISEO, "Les prises USB A / USB C sont posées en chambre 3, comme au devis, mais étaient souhaitées en chambre 1 : faut-il les déplacer ? Si oui, à quel coût ?"),
    "ch4--doubles": (ELEC, "La chambre 4 n'est pas au devis : le point lumineux et les 3 prises doubles posées (2 aux plans) sont-ils facturés en supplément ?"),
    "dressing--elec": (ISEO, "Le point lumineux et la prise du bureau / dressing figuraient sur le plan envoyé mais pas au devis : les veut-on ? Si oui, demander le prix."),
    "combles--point": (ELEC, "Le point lumineux des combles (2 directions) est hors devis : est-il facturé ?"),
    "combles--prises": (ELEC, "Aucune prise n'est au devis dans les combles, seulement 2 attentes (M1, M2) : sont-elles fermées et repérées au tableau ? Si on veut des prises, quel est le prix ?"),
}

GENERALES = [
    (ELEC, "À quelle date le contrôle de conformité par l'organisme agréé est-il prévu ? Le rapport sera-t-il transmis ?"),
    (ELEC, "Les schémas unifilaire et de position définitifs seront-ils remis, à jour de toutes les modifications ?"),
    (ELEC, "Pouvez-vous transmettre un décompte final : postes du devis non réalisés (moins-values) et suppléments, avec leurs prix ?"),
]


def main():
    doc = Document()
    s = doc.sections[0]
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = s.page_height, s.page_width
    for marge in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(s, marge, Cm(1.4))
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    total = len(QUESTIONS) + len(GENERALES)
    n_iseo = sum(1 for d, _ in QUESTIONS.values() if d == ISEO)
    doc.add_heading("Chantier Mangon – Électricité : questions ouvertes", level=0)
    p = doc.add_paragraph()
    p.add_run("Suite à la visite du 29/10/2026").bold = True
    p.add_run(f" · {total} questions : {total - n_iseo} pour l'électricien, {n_iseo} à trancher par ISEO.")
    doc.add_paragraph("Pour chaque question : ce qui a été budgété (devis), ce qui a été signé (plans) et ce qui a été "
                      "constaté sur place. La colonne « Réponse » est à compléter pendant la réunion.")
    doc.add_paragraph("Numérotation des chambres (devis) : chambre 1 = 14,12 m² (« Ch 4 » sur les plans), "
                      "chambre 2 = 16,13 m² (« Ch 3 »), chambre 3 = 16,63 m² (« Ch 2 »), chambre 4 = 11,69 m² (« Ch 1 »).")

    entetes = ["N°", "Question", "Budgété · Signé · Constaté", "Pour", "Photo", "Réponse"]
    largeurs = [0.9, 8.4, 6.2, 1.9, 2.8, 6.7]
    num = 0
    for pid, piece, alt, items in PIECES:
        lignes = [(iid, lib, dev, pl) for iid, lib, dev, pl, _ in items if f"{pid}--{iid}" in QUESTIONS]
        if not lignes:
            continue
        doc.add_heading(piece + (f" ({alt})" if alt else ""), level=2)
        t, lg = tableau(doc, entetes, largeurs)
        for iid, lib, dev, pl in lignes:
            cle = f"{pid}--{iid}"
            num += 1
            pour, question = QUESTIONS[cle]
            statut, remarque, photos = CONSTATS.get(cle, (None, "", []))
            constat = STATUTS[statut] + (f" — « {remarque} »" if remarque else "")
            contexte = f"{lib}\nDevis : {dev}\nPlans : {pl}\nConstat : {constat}"
            c = ligne(t, lg, [str(num), question, contexte, pour, None, ""])
            ecrire(c[0], str(num), gras=True)
            ecrire(c[3], pour, gras=True)
            ombrer(c[3], "DCE6F2" if pour == ELEC else "FFF2CC")
            ombrer(c[2], COUL_STATUT[statut])
            if photos:
                par = c[4].paragraphs[0]
                for ph in photos:
                    par.add_run().add_picture(str(PHOTOS / f"{ph}.jpg"), width=Cm(2.4))
            else:
                ecrire(c[4], "–")

    doc.add_heading("Questions générales", level=2)
    t, lg = tableau(doc, ["N°", "Question", "Pour", "Réponse"], [0.9, 15.6, 1.9, 8.5])
    for pour, question in GENERALES:
        num += 1
        c = ligne(t, lg, [str(num), question, pour, ""])
        ecrire(c[0], str(num), gras=True)
        ecrire(c[2], pour, gras=True)
        ombrer(c[2], "DCE6F2")

    doc.save(SORTIE)
    print(SORTIE, num, "questions")


if __name__ == "__main__":
    main()
