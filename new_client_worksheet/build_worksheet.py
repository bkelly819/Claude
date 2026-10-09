"""Builds New_Client_Worksheet.pdf, a fillable 4-page Lighthouse Private Wealth form.

Usage: python build_worksheet.py [output.pdf]
"""
import os
import sys

from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "logo.png")

NAVY = Color(.062745, .098039, .168627)
GOLD = Color(.662745, .505882, .180392)
SUBTLE = Color(.776471, .8, .847059)
MUTED = Color(.545098, .576471, .65098)
LABEL = Color(.356863, .392157, .454902)
DARK = Color(.227451, .27451, .345098)
RULE = Color(.862745, .85098, .815686)
BORDER = Color(.72549, .701961, .65098)
CARD = Color(.960784, .956863, .941176)
NOTE_BG = Color(.984314, .980392, .968627)
WHITE = Color(1, 1, 1)

LEFT, RIGHT = 46, 566
GAP = 10
ROW = 41          # label baseline to next label baseline
FIELD_H = 17

EMPLOYMENT = [" ", "Employed", "Self-Employed", "Retired", "Not Employed"]
INDUSTRY = [" ", "Financial Services", "Healthcare", "Technology", "Cybersecurity",
            "Other", "Not Applicable"]
WORTH = [" ", "Under $250,000", "$250,000 - $999,999", "$1,000,000+"]
TAX_BRACKET = [" ", "10%", "12%", "22%", "24%", "32%", "35%", "37%"]


class Form:
    def __init__(self, path):
        self.c = canvas.Canvas(path, pagesize=letter)
        self.c.setTitle("New Client Worksheet")
        self.c.setAuthor("Lighthouse Private Wealth")
        self.c.setSubject("Internal client discovery / new client worksheet")
        self.page = 0

    # ---- page chrome -------------------------------------------------------
    def start_page(self):
        c = self.c
        self.page += 1
        c.setFillColor(NAVY)
        c.rect(0, 730, 612, 62, stroke=0, fill=1)
        c.setFillColor(GOLD)
        c.rect(0, 730, 612, 3, stroke=0, fill=1)
        self.text(46, 762.05, "New Client Worksheet", "Helvetica-Bold", 15, WHITE)
        self.text(46, 748.019, "Complete in real time during the planning discovery call or meeting",
                  "Helvetica", 8.3, SUBTLE)
        self.text(566, 759.04, "Internal Use Only", "Helvetica", 8, MUTED, align="right")

    def end_page(self):
        c = self.c
        c.setStrokeColor(RULE)
        c.setLineWidth(.6)
        c.line(LEFT, 46, RIGHT, 46)
        self.text(46, 34.075, "INTERNAL USE ONLY  ·  Not for client distribution",
                  "Helvetica", 7.5, LABEL)
        self.text(306, 34.075, "New Client Worksheet", "Helvetica", 7.5, LABEL, align="center")
        self.text(566, 34.075, f"Page {self.page}", "Helvetica", 7.5, LABEL, align="right")
        c.drawImage(LOGO, 249.3566, 4, 113.2867, 27, mask="auto")
        c.showPage()

    # ---- primitives --------------------------------------------------------
    def text(self, x, y, s, font, size, color, align="left"):
        c = self.c
        c.setFont(font, size)
        c.setFillColor(color)
        {"left": c.drawString, "right": c.drawRightString,
         "center": c.drawCentredString}[align](x, y, s)

    def label(self, x, y, s):
        self.text(x, y, s, "Helvetica-Bold", 7.6, LABEL)

    def box(self, x, y, w, h, bg=WHITE):
        c = self.c
        c.setLineWidth(1)
        c.setStrokeColor(BORDER)
        c.setFillColor(bg)
        c.rect(x + .5, y + .5, w - 1, h - 1, stroke=1, fill=1)

    def field(self, name, tip, x, y, w, h=FIELD_H, options=None, multiline=False, bg=WHITE):
        self.box(x, y, w, h, bg)
        af = self.c.acroForm
        common = dict(name=name, tooltip=tip, x=x, y=y, width=w, height=h,
                      borderColor=BORDER, fillColor=bg, textColor=NAVY,
                      borderWidth=1, borderStyle="solid", fontName="Helvetica",
                      fontSize=9, forceBorder=True)
        if options:
            af.choice(options=options, value=" ", **common)
        elif multiline:
            af.textfield(fieldFlags="multiline", maxlen=None, **common)
        else:
            af.textfield(maxlen=100, **common)

    def radios(self, name, y, items):
        """items: [(x, value, label)]. Box is 10x10 with bottom at y."""
        for x, value, text in items:
            self.c.acroForm.radio(name=name, tooltip=text, value=value, selected=False,
                                  x=x, y=y, size=10, shape="square", buttonStyle="circle",
                                  borderColor=BORDER, fillColor=WHITE, textColor=NAVY,
                                  borderWidth=1, forceBorder=True)
            self.text(x + 14.5, y + 2.084, text, "Helvetica", 8.8, DARK)

    # ---- layout helpers ----------------------------------------------------
    def section(self, top, title, subtitle):
        """Returns the first label baseline below the section rule."""
        c = self.c
        c.setFillColor(GOLD)
        c.rect(46, top - 20, 3.2, 20, stroke=0, fill=1)
        self.text(58, top - 13.989, title, "Helvetica-Bold", 12.7, NAVY)
        self.text(58, top - 32.988, subtitle, "Helvetica-Oblique", 8.4, LABEL)
        c.setStrokeColor(RULE)
        c.setLineWidth(.9)
        c.line(LEFT, top - 46, RIGHT, top - 46)
        return top - 60

    def subhead(self, y, s, note=None):
        self.text(46, y, s, "Helvetica-Bold", 9.3, DARK)
        if note:
            w = self.c.stringWidth(s, "Helvetica-Bold", 9.3)
            self.text(46 + w + 7, y, note, "Helvetica-Oblique", 7.8, LABEL)
        return y - 16

    def row(self, y, cols, x0=LEFT, x1=RIGHT, gap=GAP):
        """cols: [(label, name, tip, weight, options)]. Returns next label baseline."""
        total = sum(col[3] for col in cols)
        avail = (x1 - x0) - gap * (len(cols) - 1)
        x = x0
        for lbl, name, tip, weight, options in cols:
            w = avail * weight / total
            if isinstance(lbl, tuple):  # (gold prefix, label)
                self.text(x, y, lbl[0], "Helvetica-Bold", 7.6, GOLD)
                pw = self.c.stringWidth(lbl[0] + " ", "Helvetica-Bold", 7.6)
                self.label(x + pw + 2, y, lbl[1])
            else:
                self.label(x, y, lbl)
            self.field(name, tip, x, y - 29, w, options=options)
            x += w + gap
        return y - ROW

    def save(self):
        self.c.save()


def col(label, name, tip=None, weight=1, options=None):
    return (label, name, tip or (label if isinstance(label, str) else label[1]).title(),
            weight, options)


def build(path):
    f = Form(path)

    # ===================== PAGE 1 =====================
    f.start_page()
    y = f.section(713, "1.  Client Relationship & Household",
                  "Defines the household and drives which sections below apply.")
    f.text(46, y, "HOW WILL THIS RELATIONSHIP BE REGISTERED?", "Helvetica-Bold", 8.6, LABEL)
    f.radios("relationship_type", 629, [(46, "Individual", "Individual"),
                                        (116.2, "Joint", "Joint (with Spouse / Partner)"),
                                        (260.2, "Trust", "Trust or Entity")])
    y = f.row(617.968, [col('HOUSEHOLD / RELATIONSHIP NAME (E.G. "SMITH HOUSEHOLD")',
                            "household_name", "Household / Relationship Name")])

    y = f.section(586, "2.  Primary Client Information",
                  "Identity, contact, and citizenship for the primary account holder.")
    y = f.subhead(y, "IDENTITY")
    y = f.row(509.968, [col("FIRST NAME", "primary_first", weight=177.6),
                        col("MIDDLE NAME", "primary_middle", weight=141.8),
                        col("LAST NAME", "primary_last", weight=177.6)])
    y = f.row(y, [col("DATE OF BIRTH", "primary_dob"),
                  col("MARITAL STATUS", "primary_marital",
                      options=[" ", "Single", "Married", "Divorced", "Widowed"]),
                  col("CITIZENSHIP", "primary_citizenship",
                      options=[" ", "U.S. Citizen", "Other"]),
                  col("SSN", "primary_ssn", "SSN")])
    f.subhead(428.049, "CONTACT")
    y = f.row(411.968, [col("EMAIL", "primary_email"), col("MOBILE PHONE", "primary_mobile")])
    y = f.row(y, [col("LEGAL ADDRESS (STREET, CITY, STATE, ZIP)", "primary_address",
                      "Legal Address")])
    f.text(46, 328.012, "Mailing address same as legal address?", "Helvetica-Bold", 8.4, DARK)
    f.radios("mailing_same", 321, [(220.1, "Yes", "Yes"), (268.3, "No", "No")])
    f.row(308.968, [col("MAILING ADDRESS (IF DIFFERENT)", "primary_mailing_address",
                        "Mailing Address")])

    y = f.section(276, "3.  Spouse / Joint Owner",
                  'Complete if the client is married or "Joint (with Spouse / Partner)" '
                  "was selected in Section 1.")
    y = f.row(y, [col("FIRST NAME", "joint_first", "Spouse / Joint Owner First Name"),
                  col("LAST NAME", "joint_last", "Spouse / Joint Owner Last Name"),
                  col("RELATIONSHIP TO PRIMARY", "joint_relationship",
                      options=[" ", "Spouse", "Domestic Partner", "Other"])])
    y = f.row(y, [col("DATE OF BIRTH", "joint_dob", "Spouse / Joint Owner Date of Birth"),
                  col("SSN", "joint_ssn", "Spouse / Joint Owner SSN")])
    f.row(y, [col("EMAIL ADDRESS", "joint_email", "Spouse / Joint Owner Email"),
              col("SPOUSE PHONE NUMBER", "joint_mobile", "Spouse / Joint Owner Phone Number")])
    f.end_page()

    # ===================== PAGE 2 =====================
    f.start_page()
    y = f.section(713, "4.  Beneficiaries",
                  "Allocations within each tier (Primary / Contingent) must total 100%.")
    for i in range(1, 6):
        y = f.row(y, [col(f"BENEFICIARY #{i} — FULL NAME", f"ben{i}_name",
                          f"Beneficiary {i} Full Name", 144),
                      col("RELATIONSHIP", f"ben{i}_rel", f"Beneficiary {i} Relationship", 106,
                          [" ", "Spouse", "Relative / Friend", "Non-Person (Charity/Trust)"]),
                      col("TIER", f"ben{i}_tier", f"Beneficiary {i} Tier", 81,
                          [" ", "Primary", "Contingent"]),
                      col("ALLOCATION %", f"ben{i}_pct", f"Beneficiary {i} Allocation %", 67),
                      col("PER STIRPES", f"ben{i}_stirpes", f"Beneficiary {i} Per Stirpes", 82,
                          [" ", "Yes", "No"])])

    y = f.section(455, "5.  Employment & Regulatory Disclosures",
                  "Required for AML / KYC purposes.")
    y = f.subhead(y, "PRIMARY CLIENT")
    y = f.row(y, [col("EMPLOYMENT STATUS", "employment_status", weight=111,
                      options=EMPLOYMENT),
                  col("EMPLOYER NAME", "employer_name", weight=134),
                  col("OCCUPATION", "occupation", weight=134),
                  col("INDUSTRY", "industry", weight=111, options=INDUSTRY)])
    y = f.subhead(y, "SPOUSE")
    y = f.row(y, [col("EMPLOYMENT STATUS", "spouse_employment_status",
                      "Spouse Employment Status", 111, EMPLOYMENT),
                  col("EMPLOYER NAME", "spouse_employer_name", "Spouse Employer Name", 134),
                  col("OCCUPATION", "spouse_occupation", "Spouse Occupation", 134),
                  col("INDUSTRY", "spouse_industry", "Spouse Industry", 111, INDUSTRY)])
    qy = y - 4
    f.text(46, qy, "Affiliated with a broker-dealer / FINRA member?", "Helvetica-Bold", 8.4, DARK)
    f.radios("affiliated_yn", qy - 7, [(250.9, "Yes", "Yes"), (299.1, "No", "No")])

    y = f.section(qy - 11, "6.  Financial Profile",
                  "Net worth, income, tax bracket, and current asset allocation.")
    y = f.row(y, [col("ANNUAL INCOME", "annual_income",
                      options=[" ", "Under $100,000", "$100,000 - $249,999",
                               "$250,000 - $499,999", "$500,000+"]),
                  col("NET WORTH", "net_worth", options=WORTH),
                  col("LIQUID NET WORTH", "liquid_net_worth", options=WORTH),
                  col("FEDERAL TAX BRACKET", "tax_bracket", "Federal Tax Bracket",
                      options=TAX_BRACKET)])
    y = f.subhead(y, "CURRENT ASSET ALLOCATION",
                  "(approx. % of net worth — should total 100%)")
    f.row(y, [col("CHECKING / SAVINGS %", "alloc_checking", "Checking / Savings %"),
              col("EQUITIES / STOCKS %", "alloc_equities", "Equities / Stocks %"),
              col("BONDS %", "alloc_bonds", "Bonds %"),
              col("ALTERNATIVE INVESTMENTS %", "alloc_alt", "Alternative Investments %")])
    f.end_page()

    # ===================== PAGE 3 =====================
    f.start_page()
    y = f.section(713, "7.  Accounts",
                  "Account type, value, and fee detail for each account in this relationship.")
    f.subhead(653.049, "ACCOUNT / FEE DETAIL", "(up to 5 accounts for this relationship)")
    card_top, card_h, card_gap = 647, 98, 12
    for i in range(1, 6):
        c = f.c
        c.saveState()
        c.setFillColor(CARD)
        c.setStrokeColor(BORDER)
        c.setLineWidth(.75)
        c.roundRect(46, card_top - card_h, 520, card_h, 5, stroke=1, fill=1)
        c.restoreState()
        ly = card_top - 18.132
        f.row(ly, [col((f"#{i}", "ACCOUNT TYPE"), f"acct{i}_type", f"Account {i} Type",
                       options=[" ", "Individual Brokerage", "Joint Brokerage", "Trust",
                                "Traditional IRA", "Roth IRA", "Entity / Business"]),
                   col("APPROX. ACCOUNT VALUE", f"acct{i}_value",
                       f"Account {i} Approximate Account Value"),
                   col("INVESTMENT OBJECTIVE", f"acct{i}_objective",
                       f"Account {i} Investment Objective",
                       options=[" ", "Capital Preservation", "Income", "Balanced", "Growth",
                                "Aggressive Growth"])],
              x0=60, x1=552, gap=12)
        f.row(ly - ROW, [col("MANAGED PROGRAM", f"acct{i}_program",
                             f"Account {i} Managed Program"),
                         col("MANAGER", f"acct{i}_manager", f"Account {i} Manager"),
                         col("ADVICE FEE %", f"acct{i}_fee", f"Account {i} Advice Fee %"),
                         col("$ AMOUNT", f"acct{i}_amount", f"Account {i} $ Amount")],
              x0=60, x1=552, gap=12)
        card_top -= card_h + card_gap
    f.end_page()

    # ===================== PAGE 4 =====================
    f.start_page()
    f.section(713, "8.  Notes & Additional Information",
              "Client concerns, discussion points, and additional context.")
    f.label(46, 652.968, "NOTES")
    f.field("meeting_notes", "Notes", 46, 110, 520, h=530, multiline=True, bg=NOTE_BG)
    c = f.c
    c.setStrokeColor(RULE)
    c.setLineWidth(.5)
    c.line(LEFT, 92, RIGHT, 92)
    f.row(81.968, [col("PREPARED BY (ADVISOR SIGNATURE)", "advisor_signoff",
                       "Prepared By (Advisor Signature)", 154.9),
                   col("DATE", "advisor_signoff_date", "Advisor Sign-off Date", 88.1),
                   col("REVIEWED BY (ADMIN / OPS)", "admin_signoff",
                       "Reviewed By (Admin / Ops)", 154.9),
                   col("DATE", "admin_signoff_date", "Admin Sign-off Date", 88.1)])
    f.end_page()
    f.save()


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "New_Client_Worksheet.pdf"))
