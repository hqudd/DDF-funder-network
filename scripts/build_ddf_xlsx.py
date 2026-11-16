from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

F = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=F, size=10, bold=True, color="FFFFFF")
BODY = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=14, bold=True)
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

PROGRAM_FILL = {
    "Arts": "FCE4D6",
    "Building Bridges": "E2EFDA",
    "Child Well-being": "FFF2CC",
    "Environment": "DDEBF7",
    "Medical Research": "EAD1F5",
    "Cross-program": "EDEDED",
}

STRONG = "Strong"
PARTIAL = "Partial"

# ---------------------------------------------------------------- main data
cols = [
    "DDF program", "Program / fund", "Fit with 'pooled'", "Structure", "Funders (DDF included)",
    "Administrator / host", "Amount", "DDF share", "Dates and status", "What it funds",
    "Program assignment basis", "Confidence in facts", "Notes and caveats", "Sources",
]

rows = [
    # ---------------- ARTS
    ["Arts", "Humanity AI", STRONG,
     "Pooled fund plus aligned grants from each funder's own budget",
     "Doris Duke Foundation, Ford Foundation, Lumina Foundation, Kapor Foundation, MacArthur Foundation, Mellon Foundation, Mozilla Foundation, Omidyar Network, David and Lucile Packard Foundation, Siegel Family Endowment",
     "Rockefeller Philanthropy Advisors (fiscal sponsor; manages the pooled fund)",
     "$500M over 5 years for the whole initiative. First pooled-fund round: more than $18M ($8M to 12 grantees; $10M held for an open call).",
     "Not disclosed",
     "Announced Oct 14, 2025. First pooled grants announced May 2026. Open call planned for summer 2026 (status after that not checked).",
     "Work to shape AI for people. Five issue areas: democracy, education, humanities and culture, labor and economy, security.",
     "Inferred by me (DDF has no AI program; 'humanities and culture' fits the Arts program best)",
     "High (launch); Medium (May 2026 grants)",
     "Sources disagree on the exact day of the May 2026 announcement and on the grant sizes. One aggregator lists Hewlett instead of Packard; the MacArthur release lists Packard. The split between pooled and own-budget money is not public.",
     "https://www.macfound.org/press/press-releases/humanity-ai-commits-500-million-to-build-a-people-centered-future-for-ai ; https://www.fordfoundation.org/news-and-stories/news-and-press/news/humanity-ai-announces-more-than-18-million-in-new-grants-to-shape-ai-for-the-public-good/"],

    ["Arts", "Artists Make Technology (AMT)", PARTIAL,
     "Joint multi-funder initiative led by two funders. Pooling method not stated.",
     "Doris Duke Foundation, Mozilla Foundation (lead partners); William and Flora Hewlett Foundation; Ford Foundation; MacArthur Foundation and Omidyar Network (joined after Jan 2026)",
     "Mozilla Foundation and DDF jointly",
     "$11M 3-year goal; $8M raised as of Oct 28, 2026. Plan: $6M AMT Lab, $4M AMT Pathways, $1M AMT Assemblies.",
     "Not disclosed",
     "Announced Jan 7, 2026 with $6.5M. First 20 grantees (of 392 applications) announced Oct 28, 2026. About 20 more grants possible if the goal is met.",
     "Grants and support for performing artists who build technology for their own field.",
     "Stated by DDF (DDF Arts program)",
     "High",
     "Individual funder amounts are not disclosed. Not clear whether funds go into one shared account.",
     "https://www.dorisduke.org/grants/projects/artists-make-technology ; https://www.prnewswire.com/news-releases/mozilla-foundation-and-doris-duke-foundation-back-20-artists-to-shape-what-tech-gets-built-next-302922169.html ; https://www.dorisduke.org/news/doris-duke-foundation-and-mozilla-foundation-announce-usd6-5m-commitment-to-ensure-performing-artists-are-at-the-center-of-technology-innovation"],

    ["Arts", "Mosaic Network and Fund (New York Community Trust)", STRONG,
     "Funder collaborative with a pooled grant fund",
     "19 foundations, including Doris Duke Foundation, Doris Duke Foundation for Islamic Art, Mellon, Ford, Bloomberg Philanthropies, Surdna, Howard Gilman, Booth Ferris, Lambent, Altman, Laurie M. Tisch Illumination Fund, Miranda Family Fund, New York Community Trust (list on source pages is partial)",
     "New York Community Trust (now moving to operate independently)",
     "First grant round: $4.5M to 27 arts groups (Feb 2020). DDF grant: $500K over 2 years (May 2021).",
     "$500K over 2 years (one 2021 grant)",
     "Set up 2018 as a five-year collaborative. Grants made from 2020. Now an independent membership group with a new executive director (date not checked).",
     "General operating grants to arts groups led by African, Latinx, Asian, Arab and Native American people in New York City, plus a learning network.",
     "Stated by DDF (Arts)",
     "High",
     "DDF co-chaired the advisory committee and co-commissioned the 2017 research that led to the fund.",
     "https://thenytrust.org/collaborative-fund/mosaic-network-and-fund ; https://dorisduke.org/grants/what-weve-funded/Grant-Recipients/community-funds-inc ; https://philanthropynewyork.org/news/mosaic-network-and-fund-new-york-community-trust-grants-45-million-arts-nonprofits-color-new"],

    ["Arts", "Artist Relief (COVID-19 emergency fund)", STRONG,
     "Coalition emergency fund with one shared pool",
     "More than 20 arts funders. Mellon gave $5M seed funding. Doris Duke Charitable Foundation and Ford are named donors.",
     "Seven arts groups: Academy of American Poets, Artadia, Creative Capital, Foundation for Contemporary Arts, MAP Fund, National YoungArts Foundation, United States Artists",
     "$10M raised in about 10 days (Apr 2020); $5,000 grants",
     "Not disclosed",
     "Launched Apr 2020. Current status not checked.",
     "Emergency $5,000 grants to individual artists in all fields.",
     "Stated (DDF is a performing-arts funder)",
     "Medium",
     "Found through press coverage only. I did not find the DDF amount.",
     "https://hyperallergic.com/artist-relief.md ; https://www.artforum.com/news/covid-19-relief-fund-to-give-10-million-to-artists-82702"],

    ["Arts", "National Dance Project (NEFA)", PARTIAL,
     "Multi-funder grant program run by an intermediary",
     "Mellon Foundation (lead funder until 2025) and Doris Duke Foundation (lead funder). Other funders not checked.",
     "New England Foundation for the Arts (NEFA)",
     "Not found",
     "Not disclosed",
     "Mellon ended its support after 2025. DDF says it plans to continue. Details of the final cycle were to be shared later.",
     "Production and touring grants for dance.",
     "Stated (DDF Arts)",
     "Medium",
     "Included because two lead funders fund one program. Funds may not be formally pooled.",
     "https://www.kunc.org/2025-09-09/what-will-happen-to-dance-with-a-major-grant-program-changing ; https://www.npr.org/transcripts/nx-s1-5526923"],

    ["Arts", "MAP Fund national regranting program", PARTIAL,
     "Multi-funder regranting program run by an intermediary",
     "Doris Duke Foundation (primary funder since 2008), Mellon Foundation (since 2010), Jerome Foundation, Howard Gilman Foundation, Walder Foundation, public funders",
     "MAP Fund",
     "DDF: $21.8M since 2008. Mellon: $6.8M since 2010.",
     "$21.8M since 2008 (about $1.43M a year at the end)",
     "Ended. DDF and Mellon stopped funding regranting through MAP; announced Dec 12, 2024.",
     "Grants to performing artists for new work.",
     "Stated (DDF Arts)",
     "High",
     "Included for completeness. Program has ended.",
     "https://www.mapfund.org/news/important-announcement-regarding-map-fund-grant-program ; https://www.americantheatre.org/2024/12/13/map-fund-retires-national-grant-program/"],

    # ---------------- BUILDING BRIDGES
    ["Building Bridges", "None found", "n/a",
     "n/a", "n/a", "n/a", "n/a", "n/a", "n/a",
     "DDF's Building Bridges grants that I found (Sundance Institute $1M, Center for Asian American Media fund, Islamic Scholarship Fund $500K) come from DDF alone.",
     "n/a", "n/a",
     "Absence of evidence only. Not proof that none exists. See 'Leads and gaps'.",
     "https://www.sundance.org/blogs/doris-duke-foundation-announces-1m-grant-to-sundance-institute-expanding-commitment-to-uplifting-muslim-storytelling-with-new-building-bridges-fellowship-and-completion-fund/ ; https://documentary.org/node/41073"],

    # ---------------- CHILD WELL-BEING
    ["Child Well-being", "OPT-In for Families (Opportunities for Prevention & Transformation)", PARTIAL,
     "DDF-led initiative with outside co-funders. Pooling method not stated.",
     "Doris Duke Foundation (lead), The Duke Endowment, Aviv Foundation",
     "Doris Duke Foundation",
     "$33M over 3 years; about $9M for each of four sites",
     "Up to $3M a site in direct family support plus technical help (the rest of each site's $9M is not itemized)",
     "Launched Mar 11, 2024. Sites: Kentucky, Oregon, South Carolina, Washington DC.",
     "Prevention-focused child well-being systems.",
     "Stated by DDF (Child Well-being)",
     "Medium",
     "Press says a 'coalition of funders'. Co-funder amounts are not public.",
     "https://dorisduke.org/news/doris-duke-foundation-launches-new-opt-in-for-families-initiative-to-refocus-the-child-welfare-system-on-prevention ; https://www.insidephilanthropy.com/home/2024/4/15/the-doris-duke-charitable-foundation-stands-up-an-initiative-to-transform-the-child-welfare-system"],

    ["Child Well-being", "Think of Us - Audacious Project award", STRONG,
     "Pooled award through the Audacious Project (housed at TED), topped up by named funders",
     "Audacious Project donor pool; Annie E. Casey Foundation; Conrad N. Hilton Foundation; Doris Duke Foundation; Aviv Foundation",
     "The Audacious Project (TED)",
     "$47.5M over 5 years",
     "Not disclosed",
     "Announced Apr 17, 2023.",
     "Data and software to change child welfare and foster care systems.",
     "Inferred by me (child welfare is a Child Well-being topic)",
     "High",
     "DDF has also made separate grants to Think of Us. Those are not part of this count.",
     "https://www.philanthropy.com/news/foster-care-veteran-gets-47-5-million-from-audacious-project-to-redesign-child-welfare-system ; https://www.citybiz.co/article/409344/think-of-us-raised-47-5m-for-tech-to-improve-the-foster-system/ ; https://aecf.org/blog/think-of-us-to-expand-its-efforts-with-investment-from-the-audacious-project"],

    # ---------------- ENVIRONMENT
    ["Environment", "Tribal Nations Conservation Pledge", PARTIAL,
     "Coordinated multi-funder pledge. Each funder commits its own money.",
     "15 funders, including Doris Duke Foundation, Alaska Conservation Foundation, Christensen Fund, Decolonizing Wealth Project, Gordon and Betty Moore Foundation, Hewlett Foundation, Wyss Foundation",
     "Native Americans in Philanthropy and the Biodiversity Funders Group (organizers)",
     "$102.5M over 5 years",
     "$32M (about one third)",
     "Announced Mar 21, 2023 at a White House summit.",
     "Tribal and Indigenous-led conservation. DDF's share runs through its own Indigenous-led conservation initiative.",
     "Stated by DDF (Environment)",
     "High",
     "DDF's release spells it 'Tribal National Conservation Pledge'; other sources say 'Tribal Nations'. Not a single shared pool.",
     "https://www.dorisduke.org/news/doris-duke-foundation-and-partners-announce-over-100-million-for-tribal-led-conservation ; https://news.mongabay.com/2023/04/funders-commit-102-5-million-to-support-tribal-led-conservation-efforts-in-the-u-s/ ; https://www.christensenfund.org/?p=17081"],

    ["Environment", "Catalyst Fund (Network for Landscape Conservation)", PARTIAL,
     "Joint fund from two funders, run by a third party",
     "Doris Duke Foundation, William and Flora Hewlett Foundation",
     "Network for Landscape Conservation (fiscally sponsored by the Center for Large Landscape Conservation)",
     "Grants up to $25,000 over 2 years. About $2.5M awarded in total through the 2025 cycle.",
     "Not disclosed",
     "Launched 2019. 2025 cycle: 18 grants (proposals due May 16, 2025). Later cycles not checked.",
     "Capacity grants for local landscape conservation partnerships. Part is reserved for Tribal-led partnerships.",
     "Stated by DDF (Environment)",
     "High",
     "103 partnerships funded so far, 36 of them Indigenous-led.",
     "https://landscapeconservation.org/knowledge-center/stories/the-catalyst-fund-shows-a-way-to-the-successful-conservation-stewardship-of-americas-landscapes/ ; https://landscapeconservation.org/wp-content/uploads/2017/12/Catalyst-Fund_RFP_2025-.pdf"],

    # ---------------- MEDICAL RESEARCH
    ["Medical Research", "Pathways for Health Research Collective (More Than Medicine)", STRONG,
     "Funder collective. Includes a one-time collaborative request for applications.",
     "Doris Duke Foundation (lead), American Cancer Society, American Heart Association, Burroughs Wellcome Fund, Dana Foundation, Donaghue Foundation, Robertson Foundation, Susan G. Komen, and other partners",
     "Doris Duke Foundation leads",
     "Not found. Awards 'commensurate with project scope'.",
     "Not disclosed",
     "Collective launched Sept 2024. Symposia in spring 2025. Open now: pre-proposals due Nov 18, 2026, 3:00 p.m. ET.",
     "Research on health solutions science: how care delivery, policy and living conditions shape health.",
     "Stated by DDF (Medical Research)",
     "High",
     "Deadline is 8 days from today (Nov 10, 2026). Earlier name: Collective to Strengthen Pathways for Health Research.",
     "https://morethanmedicine.org/ ; https://dana.org/article/philanthropic-collective-aims-to-improve-disease-prevention-and-care/ ; https://dorisduke.org/news--insights/articles/revisiting-research-pathways"],

    ["Medical Research", "COVID-19 Fund to Retain Clinical Scientists", STRONG,
     "One-time collaborative funding competition",
     "Doris Duke Charitable Foundation, American Heart Association, Burroughs Wellcome Fund, John Templeton Foundation, Rita Allen Foundation, Walder Foundation",
     "Doris Duke Foundation organized it",
     "$12.1M to 22 medical schools (2 years)",
     "Not disclosed",
     "Awards announced Nov 17, 2021. Closed. DDF sees no new grants soon.",
     "Funds for early-career physician-scientists with family caregiving duties.",
     "Stated by DDF (Medical Research)",
     "High",
     "Builds on DDF's own Fund to Retain Clinical Scientists (2015).",
     "https://dorisduke.org/funding-areas/medical-research/fund-to-retain-clinical-scientists ; https://www.dorisduke.org/siteassets/news--insights/2021-press-releases/11.16.21-covid19-frcs-press-release.pdf"],

    # ---------------- CROSS-PROGRAM
    ["Cross-program", "NYC COVID-19 Response & Impact Fund", STRONG,
     "Pooled emergency fund of New York City funders",
     "Doris Duke Charitable Foundation and other New York City area funders (full list not checked)",
     "New York Community Trust (per the Mosaic source; not confirmed on a primary page)",
     "$75M",
     "Not disclosed",
     "Announced 2020. Current status not checked.",
     "Support for service providers in health, hunger and the arts.",
     "Inferred by me (spans arts and health)",
     "Medium",
     "Two secondary sources only. Verify before you cite.",
     "https://www.insidephilanthropy.com/grants-d/tag/Muslim ; https://china.db.com/usa/documents/responsibility/publications/2020/Promising_Participatory_Model_for_Diversifying_Arts_Funding.pdf"],
]

# ---------------------------------------------------------------- related (not pooled)
rel_cols = ["DDF program", "Item", "Why it is not on the main list", "Funders", "Date", "Sources"]
rel_rows = [
    ["Arts", "DDF and Ford technology infrastructure grants",
     "Five grants made 'in partnership with' Ford. Co-funded grants. No pool is described.",
     "Doris Duke Foundation, Ford Foundation", "Jul 6, 2026; more than $4M over 5 years",
     "https://www.dorisduke.org/news/doris-duke-foundation-announces-new-grants-supporting-technology-infrastructure-in-the-performing-arts"],
    ["Arts", "Five-foundation $1.7B pandemic commitment",
     "Headline says the foundations would 'activate' more than $1.7B in enhanced grantmaking. I read only the headline and caption. It looks like each foundation raising its own payout, not a pool. Terms not verified.",
     "Mellon, Doris Duke Charitable Foundation, Ford, W.K. Kellogg, MacArthur", "Jun 2020",
     "https://www.artsjournal.org/2020/06/11/five-foundations-step-up-with-1-7-billion-to-help-non-profit-arts-humanities/"],
    ["Arts", "Doris Duke Foundation Performing Artist Recovery Fund",
     "DDF-only money, held at New York Community Trust.",
     "Doris Duke Foundation", "Oct 2021; $3M to about 300 artists",
     "https://philanthropynewyork.org/news/partnership-between-doris-duke-charitable-foundation-and-new-york-community-trust-awards-3"],
    ["Building Bridges", "Sundance Building Bridges Fellowship and Completion Fund; CAAM U.S. Muslim Documentary Fund; Islamic Scholarship Fund grant",
     "DDF is the only funder named in the sources I read.",
     "Doris Duke Foundation", "2024-2025",
     "https://www.sundance.org/blogs/doris-duke-foundation-announces-1m-grant-to-sundance-institute-expanding-commitment-to-uplifting-muslim-storytelling-with-new-building-bridges-fellowship-and-completion-fund/ ; https://documentary.org/node/41073"],
    ["Environment", "Biodiversity Funders Group membership",
     "A network of funders that shares learning. It is not a grant fund.",
     "Doris Duke Foundation and other members", "Ongoing",
     "https://www.sourcewatch.org/index.php/The_Consultative_Group_on_Biological_Diversity"],
    ["Medical Research", "Health Research Alliance data sharing",
     "DDF is one of 33 funders that share grant data. No money is pooled.",
     "Doris Duke Foundation and 32 other funders", "Ongoing",
     "https://www.dorisduke.org/news--insights/articles/revisiting-research-pathways"],
]

# ---------------------------------------------------------------- leads and gaps
lead_cols = ["Type", "Item", "What I found", "Next step"]
lead_rows = [
    ["Lead (Environment)", "New collaborative to reverse biodiversity loss in the U.S.",
     "DDF gave $250,000 over 1 year to the Consultative Group on Biological Diversity (Biodiversity Funders Group) on Nov 11, 2024 'to support the development of a new collaborative philanthropic effort'. I found no launch announcement.",
     "Search Biodiversity Funders Group news. Ask DDF's Environment program."],
    ["Lead (Environment)", "Conservation Loan Collaborative",
     "DDF gave $250,000 (2024) to launch a platform that manages revolving loan funds for conservation groups, coalitions and funders. Funders may add money to it.",
     "Read the DDF grant record and the grantee site to see who else funds it."],
    ["Lead (Environment)", "Funder learning and coordination for Indigenous-led conservation",
     "A $3.6M DDF grant (2024) 'to accelerate funder learning, coordination and investment in Indigenous-led conservation'. The grantee name was cut off in my results.",
     "Open DDF's grants database and find this grant."],
    ["Gap", "DDF grants database not searched in full",
     "DDF's grants pages load by script. I only saw search snippets of individual grant records.",
     "Export DDF's grant list, or pull the foundation's records from Candid. Search for words like 'collaborative', 'pooled', 'fund' and 'partnership'."],
    ["Gap", "Source quality",
     "Several Humanity AI and AMT details come from aggregator sites. I used them only to fill gaps around primary pages from MacArthur, Ford, DDF and Mozilla.",
     "Confirm the 'Medium' rows with the funders before you publish."],
    ["Gap", "Not covered",
     "Duke Farms and Shangri La (DDF's two centers), older programs before 2015, and funder collaboratives where DDF is a member but not a named funder.",
     "Tell me if you want any of these added."],
]


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = BORDER


def write_table(ws, headers, data, widths, program_col=None):
    ws.append(headers)
    style_header(ws, len(headers))
    for r in data:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY
            cell.alignment = WRAP
            cell.border = BORDER
        if program_col is not None:
            pc = row[program_col]
            pc.font = BOLD
            pc.fill = PatternFill("solid", fgColor=PROGRAM_FILL.get(pc.value, "FFFFFF"))
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2" if program_col is not None else "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    ws.row_dimensions[1].height = 32


wb = Workbook()

# ---- Summary sheet (formulas count the main sheet)
ws = wb.active
ws.title = "Summary"
ws["A1"] = "Doris Duke Foundation: multi-funder grant programs"
ws["A1"].font = TITLE
ws["A2"] = "Research date: Nov 10, 2026. Sources: dorisduke.org and the wider web. Details are in the 'Programs' tab."
ws["A2"].font = BODY

ws["A4"] = "DDF program"
ws["B4"] = "Programs found"
ws["C4"] = "Strong fit"
ws["D4"] = "Partial fit"
style_header(ws, 4, row=4)
programs = ["Arts", "Building Bridges", "Child Well-being", "Environment", "Medical Research", "Cross-program"]
for i, p in enumerate(programs):
    r = 5 + i
    ws.cell(row=r, column=1, value=p)
    ws.cell(row=r, column=2, value=f'=COUNTIFS(Programs!$A:$A,$A{r},Programs!$C:$C,"Strong")+COUNTIFS(Programs!$A:$A,$A{r},Programs!$C:$C,"Partial")')
    ws.cell(row=r, column=3, value=f'=COUNTIFS(Programs!$A:$A,$A{r},Programs!$C:$C,"Strong")')
    ws.cell(row=r, column=4, value=f'=COUNTIFS(Programs!$A:$A,$A{r},Programs!$C:$C,"Partial")')
    for c in range(1, 5):
        cell = ws.cell(row=r, column=c)
        cell.font = BOLD if c == 1 else BODY
        cell.border = BORDER
        if c == 1:
            cell.fill = PatternFill("solid", fgColor=PROGRAM_FILL[p])
tr = 5 + len(programs)
ws.cell(row=tr, column=1, value="Total")
for c, L in zip((2, 3, 4), "BCD"):
    ws.cell(row=tr, column=c, value=f"=SUM({L}5:{L}{tr-1})")
for c in range(1, 5):
    cell = ws.cell(row=tr, column=c)
    cell.font = BOLD
    cell.border = BORDER

notes = [
    "How to read this file",
    "Strong fit: sources call it a pooled fund, a collaborative fund or a joint competition.",
    "Partial fit: several funders back one effort, but the sources do not say the money is pooled.",
    "Programs are grouped by DDF's five national programs. 'Cross-program' holds items that span programs.",
    "The 'Program assignment basis' column shows where I chose the program myself. Check those rows.",
    "Confidence 'High' means a primary page (funder or partner) confirms the facts. 'Medium' means secondary sources only.",
    "Building Bridges: I found no multi-funder program. That does not prove none exists.",
    "Open now: Pathways for Health Research Collective pre-proposals are due Nov 18, 2026, 3:00 p.m. ET.",
    "Tabs: 'Related, not pooled' lists items I checked and left out. 'Leads and gaps' lists items to follow up.",
]
for i, t in enumerate(notes):
    c = ws.cell(row=tr + 2 + i, column=1, value=t)
    c.font = BOLD if i == 0 else BODY
    c.alignment = Alignment(wrap_text=False, vertical="top")
ws.column_dimensions["A"].width = 26
for L in "BCD":
    ws.column_dimensions[L].width = 16

# ---- Programs sheet
wp = wb.create_sheet("Programs")
write_table(wp, cols, rows,
            [16, 32, 12, 30, 48, 32, 36, 24, 36, 40, 28, 18, 44, 60], program_col=0)

# ---- Related sheet
wr = wb.create_sheet("Related, not pooled")
write_table(wr, rel_cols, rel_rows, [16, 40, 60, 36, 28, 60], program_col=0)

# ---- Leads sheet
wl = wb.create_sheet("Leads and gaps")
write_table(wl, lead_cols, lead_rows, [22, 42, 80, 60])

for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False

wb.save("DDF_Pooled_Fund_Programs.xlsx")
print("saved")
