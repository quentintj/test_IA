"""Génère la note de préparation de la réunion avec l'électricien (Word puis PDF)."""
from docx import Document
from docx.shared import Cm, Pt, RGBColor

from generer_bilan import ICI, ecrire, ligne, ombrer, tableau
from generer_questions import numeros

SORTIE = ICI / "Reunion_electricien_Mangon_preparation.docx"
Q = numeros()


def q(*cles):
    return "Q" + ", Q".join(str(Q[c]) for c in cles)


OBJECTIF = ("Repartir avec, par écrit : la liste des corrections et leurs dates, et un décompte final "
            "(moins-values et suppléments) accepté des deux côtés.")

AVANT = [
    "Décider en interne des 6 questions « ISEO » du PDF des questions : chambre de la commande depuis la tête de lit "
    f"et des prises USB ({q('ch1--doubles', 'ch3--usb')}), éclairage côté jardin ({q('garage--jardin')}), "
    f"bureau / dressing ({q('dressing--elec')}), prise sous la fenêtre de la cuisine ({q('cuisine--pdt')}), "
    f"erreur exacte sur les interrupteurs de la salle à manger ({q('sam--inter')}).",
    "Avoir sous la main : le devis, les plans électriques signés, le bilan du 29/10, les photos et le PDF des questions.",
    "Fixer ce qu'on accepte de payer en supplément au maximum, et ce sur quoi on ne cède pas (voir points 4 et 6).",
]

OUVERTURE = [
    "« On a fait le tour complet le 29/10, pièce par pièce, avec le devis et les plans signés en main. "
    "Une grande partie est conforme (48 points sur 86), merci pour ça. »",
    "« Il reste des points prévus au devis ou aux plans qui ne sont pas faits ou pas conformes. On veut les "
    "passer un par un et convenir ensemble d'une date pour chaque correction. »",
    "« On veut aussi clarifier les écarts par rapport au devis, dans les deux sens, pour arriver à un décompte "
    "final clair. On préfère tout régler aujourd'hui plutôt qu'à la facture. »",
]

# (point, base, ce qu'on attend, réf.)
EXIGE = [
    ("Appliques murales du salon (2, en 2 directions)", "Devis + plans (A3)", "Pose", q("salon--appl")),
    ("Point lumineux sous meuble de la cuisine", "Devis", "Pose", q("cuisine--meuble")),
    ("Alimentations pompe à chaleur (int. + ext.) et chauffe-eau", "Devis (3)", "Réalisation conforme", q("garage--pac")),
    ("Prises simples de la chambre 3", "Devis (3) + plans (2)", "Pose", q("ch3--simples")),
    ("3e prise double de la chambre 2", "Devis + plans (3)", "Pose, ou explication si elle existe", q("ch2--doubles")),
    ("Prise double indépendante de la salle de douche", "Devis", "Recâblage : ne doit pas dépendre de la lumière", q("douche--douche")),
    ("Interrupteur des spots de la cuisine", "Devis (déplacé à la demande d'ISEO)", "Déplacement au bon endroit", q("cuisine--inter")),
    ("Témoins des appliques extérieures (buanderie)", "Devis (1 direction + témoin)", "Ajout du témoin", q("buanderie--appl")),
    ("Applique du hall des chambres 3-4 (2 directions + témoin)", "Devis", "Applique, 2 directions et témoin", q("hall34--appl")),
    ("Prises NIKO Hydro (extérieur et garage)", "Devis : noires", "Remplacement par des noires", q("buanderie--ext", "garage--prise")),
    ("Tube LED du garage", "Devis + plans", "Remonter au plus haut", q("garage--tube")),
    ("Témoins des commandes de l'escalier", "Devis", "Vérifier et compléter", q("hall12--commandes")),
    ("Liaisons équipotentielles", "Devis", "Réalisées et prouvées par le contrôle", q("sdb--equipot", "douche--equipot", "tableau--equipot")),
]

MOINS_VALUES = [
    ("Spots des halls", "8 au devis, 6 posés", "2 spots à déduire", q("halls--spots")),
    ("Spots du garde-manger", "3 au devis, 2 posés", "1 spot à poser ou à déduire", q("garde--spots")),
    ("Différentiels et disjoncteurs", "Tétrapolaires au devis, bipolaires posés ; ni 25 A ni 6 A", "Justification et différence de prix à déduire", q("tableau--diff")),
    ("Points lumineux de la salle de douche", "2 au devis, 1 aux plans", "Si un seul est posé : à déduire ou à poser", q("douche--points")),
]

# (point, situation, position)
SUPPLEMENTS = [
    ("2 spots en plus dans la cuisine", "Aux plans signés, pas au devis", "Dû, au prix unitaire du devis", q("cuisine--spots")),
    ("Chambre 4 : point lumineux + 3 prises doubles", "Aux plans (2 prises), pas au devis", "Point + 2 prises au prix du devis ; la 3e prise à discuter", q("ch4--doubles")),
    ("Tableau du garage", "Aux plans, pas au devis", "Dû si on l'a demandé, prix détaillé exigé", q("garage--td")),
    ("Alimentation de la chaudière", "Aux plans, pas au devis, non conforme", "Rien à payer tant qu'elle n'est pas conforme", q("garage--chaud")),
    ("Spots au lieu de points classiques (escalier, hall ch. 1-2)", "Devis : points ; plans : spots", "5 spots comme convenu, sans surcoût", q("hall12--eclairage")),
    ("Détecteur du garde-manger", "Plan envoyé, pas au devis", "Monnaie d'échange : à lâcher si le reste est réglé", q("garde--detect")),
    ("Point lumineux des combles", "Plans, pas au devis", "Dû au prix du devis s'il a été demandé", q("combles--point")),
]

CEDER = [
    "Accepter les suppléments qui figurent aux plans signés, au prix unitaire du devis, en échange de corrections "
    "rapides et gratuites de tous les manquements.",
    "Laisser les spots dans l'escalier et le hall des chambres 1-2 au lieu des points classiques, à condition qu'il n'y ait pas de surcoût.",
    "Ne pas réclamer de moins-value sur les petits écarts (détecteur du garde-manger, une prise en plus en chambre 4) "
    "si le reste est réglé sans discussion.",
]

NE_PAS_CEDER = [
    "Tout ce qui est au devis et pas fait (ou mal fait) est corrigé sans frais.",
    "Aucun supplément n'est dû s'il n'est ni au devis ni aux plans signés et qu'il n'a pas été commandé par écrit.",
    "Rien de neuf n'est accepté oralement : chaque supplément ou modification fait l'objet d'un prix écrit avant exécution.",
    "Le solde n'est payé qu'une fois les réserves levées et le rapport de contrôle conforme reçu.",
]

RESULTAT = [
    "Une date de retour sur chantier pour les corrections, et qui vient.",
    "La date du contrôle de conformité par l'organisme agréé, et l'engagement de transmettre le rapport.",
    "Les schémas unifilaire et de position définitifs et un listing du tableau à la numérotation des chambres du devis.",
    "Le PV de mesure de la terre.",
    "Un décompte final écrit : moins-values et suppléments, poste par poste.",
    "Un compte rendu envoyé par mail sous 48 h, reprenant chaque point décidé, avec les réponses notées dans le PDF des questions.",
]

PIEGES = [
    "La numérotation des chambres diffère entre le devis, les plans et le tableau. Toujours désigner une chambre par sa "
    "surface (14,12 / 16,13 / 16,63 / 11,69 m²), pas par son numéro.",
    "Ne pas valider de supplément sur place. Répondre : « Envoyez-nous un prix, on revient vers vous. »",
    "S'il conteste un manquement, montrer la ligne du devis et la photo, sans débattre au-delà.",
    "Garder le ton factuel : le but est de finir le chantier, pas de régler des comptes.",
]


def puces(doc, items, style="List Bullet"):
    for it in items:
        doc.add_paragraph(it, style=style)


def table_points(doc, entetes, largeurs, lignes, couleur):
    t, lg = tableau(doc, entetes, largeurs)
    for valeurs in lignes:
        c = ligne(t, lg, list(valeurs))
        ecrire(c[0], valeurs[0], gras=True)
        ombrer(c[0], couleur)


def main():
    doc = Document()
    s = doc.sections[0]
    for marge in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(s, marge, Cm(1.6))
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    doc.add_heading("Réunion avec l'électricien – Chantier Mangon", level=0)
    p = doc.add_paragraph()
    p.add_run("Ce qu'on dit, ce qu'on exige, ce qu'on peut négocier, ce qu'on doit obtenir").bold = True
    note = doc.add_paragraph().add_run(
        "Les numéros Q renvoient au PDF des questions ouvertes, qui sert de fil conducteur pendant la réunion.")
    note.italic = True
    note.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

    doc.add_heading("1. Objectif", level=1)
    doc.add_paragraph().add_run(OBJECTIF).bold = True

    doc.add_heading("2. Avant la réunion (en interne)", level=1)
    puces(doc, AVANT)

    doc.add_heading("3. Ouverture : ce qu'on dit", level=1)
    puces(doc, OUVERTURE)

    doc.add_heading("4. Ce qu'on exige : à corriger sans frais", level=1)
    doc.add_paragraph("Ces points sont prévus au devis et/ou aux plans signés. Ce n'est pas négociable : on attend une date pour chacun.")
    table_points(doc, ["Point", "Base", "Ce qu'on attend", "Réf."], [6.2, 3.8, 5.2, 2.6], EXIGE, "F4CCCC")

    doc.add_heading("5. Ce qu'on réclame en moins-value", level=1)
    table_points(doc, ["Point", "Constat", "Ce qu'on demande", "Réf."], [4.6, 5.2, 5.4, 2.6], MOINS_VALUES, "DCE6F2")

    doc.add_heading("6. Les suppléments qu'il va présenter : notre position", level=1)
    doc.add_paragraph("Règle : s'il figure aux plans signés, il est probablement dû, au prix unitaire du devis. "
                      "S'il n'est ni au devis ni aux plans et n'a pas été commandé par écrit, il n'est pas dû.")
    table_points(doc, ["Supplément", "Situation", "Notre position", "Réf."], [5.0, 4.4, 5.8, 2.6], SUPPLEMENTS, "FFF2CC")

    doc.add_heading("7. Marge de négociation", level=1)
    p = doc.add_paragraph()
    p.add_run("On peut lâcher :").bold = True
    puces(doc, CEDER)
    p = doc.add_paragraph()
    p.add_run("On ne lâche pas :").bold = True
    puces(doc, NE_PAS_CEDER)

    doc.add_heading("8. Ce qu'on doit obtenir en sortant", level=1)
    puces(doc, RESULTAT, style="List Number")

    doc.add_heading("9. Points d'attention", level=1)
    puces(doc, PIEGES)

    doc.save(SORTIE)
    print(SORTIE)


if __name__ == "__main__":
    main()
