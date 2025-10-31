from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
from docx.shared import Pt, Cm
from docx.oxml.ns import qn as qn_ns


def set_default_style(document):
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)


def set_margins(section):
    # Common academic margins (RU): left 3 cm, right 1.5 cm, top/bottom 2 cm
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)


def add_paragraph(document, text, bold=False, align=None, spacing=1.5):
    p = document.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    if align is not None:
        p.alignment = align
    # Set 1.5 line spacing
    p.paragraph_format.line_spacing = spacing
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    return p


def add_heading(document, text):
    p = add_paragraph(document, text, bold=True)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    return p


def add_title_page(document):
    # Simple title page without page number
    add_paragraph(document, 'Реферат', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph(document, 'Организация закупок в сфере IT (2025)', bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
    document.add_paragraph()  # spacer
    add_paragraph(document, 'Дисциплина: Управление ИТ / ИТ-менеджмент', align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph(document, 'Студент: 4 курс, 20 лет', align=WD_ALIGN_PARAGRAPH.CENTER)
    add_paragraph(document, 'Город, 2025', align=WD_ALIGN_PARAGRAPH.CENTER)

    # Add section break to start next page
    document.add_page_break()


def add_numbering_footer(section):
    # Adds page number field in the footer (PAGE)
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Create field for current page number
    fld_simple = OxmlElement('w:fldSimple')
    fld_simple.set(qn('w:instr'), 'PAGE')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rPr.append(rFonts)
    r.append(rPr)
    fld_simple.append(r)
    p._p.append(fld_simple)


def build_document():
    document = Document()

    # Defaults and margins
    set_default_style(document)
    for section in document.sections:
        set_margins(section)

    # Title page
    add_title_page(document)

    # Add page numbers starting this section
    for section in document.sections:
        add_numbering_footer(section)

    # Body content
    add_paragraph(document, 'Введение', bold=True)
    add_paragraph(document, (
        'Организация закупок в IT в 2025 году — это уже не просто «купить компьютеры и софт». '
        'Почти всё уехало в облака, сервисы стали подписками, в работу активно вошли модели ИИ, '
        'а требования к безопасности и устойчивости заметно выросли. Неправильно выстроенный процесс '
        'ведёт к vendor lock-in, переплатам, проблемам с комплаенсом и уязвимостям в цепочке поставок. '
        'В этом реферате я собрал практичный взгляд на то, как в 2025 году организовывать IT‑закупки — '
        'от жизненного цикла до контрактов и метрик.'
    ))

    add_heading(document, '1. Что такое IT‑закупки и чем они отличаются')
    add_paragraph(document, (
        'IT‑закупки — это приобретение программного обеспечения (SaaS, on‑prem), облачных ресурсов (IaaS/PaaS), '
        'лицензий, оборудования (серверы, сети, ноутбуки), услуг интеграторов и поддержки. Ключевые отличия: '
        'подписки вместо «вечных» лицензий, повышенная значимость безопасности и приватности, быстрые циклы обновлений, '
        'архитектурные последствия выбора и акцент на совокупной стоимости владения (TCO), а не только на стартовой цене.'
    ))

    add_heading(document, '2. Жизненный цикл закупки')
    add_paragraph(document, (
        'Инициация потребности → бизнес‑кейс → выбор процедуры (RFI/RFP/RFQ, тендер, каталог) → '
        'описание требований (функциональные и нефункциональные) → оценка поставщиков → договор и согласования → '
        'пилот/POC → внедрение → приёмка → эксплуатация и пересмотр. На каждом шаге фиксируются измеримые критерии '
        'и риски, отдельно — требования к данным и безопасности.'
    ))

    add_heading(document, '3. Модели организации')
    add_paragraph(document, (
        'Централизованная модель даёт контроль и скидки, но медленнее. Децентрализованная — быстрее и ближе к нуждам '
        'команд, но рискует разнобоем и дублированием. В 2025 наиболее рабочий вариант — гибрид: центр компетенций '
        'задаёт стандарты, каталоги «разрешённых» решений и проводит сложные сделки; продуктовые команды делают '
        'типовые закупки быстро внутри рамок.'
    ))

    add_heading(document, '4. Методы закупок')
    add_paragraph(document, (
        'RFI — чтобы понять рынок; RFP — когда нужна проработка архитектуры и безопасности; RFQ — при чётких спецификациях. '
        'Тендеры и аукционы повышают прозрачность, но иногда чрезмерно давят на цену в ущерб качеству. Каталоги и '
        'рамочные соглашения ускоряют типовые покупки. Маркеты облаков (AWS/Azure/GCP) упрощают оформление, но повышают риск lock‑in.'
    ))

    add_heading(document, '5. Специфика 2025 года: облака, AI, безопасность, устойчивость')
    add_paragraph(document, (
        'Облака и FinOps стали «по умолчанию»: важнее предсказуемость и управляемость затрат, чем низкая стартовая цена. '
        'Для ИИ‑закупок действуют требования EU AI Act: проверка источников данных, управление рисками и объяснимость. '
        'Цепочки поставок требуют SBOM и практик безопасной разработки. Важны локация данных и юрисдикции, '
        'а также устойчивость (ESG): энергоэффективность, углеродный след, отчётность.'
    ))

    add_heading(document, '6. Оценка поставщиков и рисков')
    add_paragraph(document, (
        'Безопасность и комплаенс: ISO/IEC 27001:2022, SOC 2 Type II, шифрование, SSO/MFA, Zero Trust, отчёты аудиторов, SBOM. '
        'Лицензии и IP: права на данные и модели, совместимость OSS‑лицензий, отсутствие «тёмных» зависимостей. '
        'Экономика: TCO (подписка, хранение, egress, администрирование, интеграции, обучение, миграции, SLA‑кредиты) и ROI. '
        'Lock‑in: экспорт данных в открытых форматах, API‑доступ, план выхода и оценка стоимости миграции.'
    ))

    add_heading(document, '7. Контракты, SLA и SLO')
    add_paragraph(document, (
        'SLA: доступность, время реакции и восстановления, окна обновлений, поддержка 24/7, RTO/RPO. '
        'SLO и SLA‑кредиты — пороговые значения и механизм компенсаций. DPA — роли контролёра/процессора, '
        'местоположение данных и подобработчики. Для ИИ — запрет дообучения на клиентских данных, контроль датасетов, '
        'требования к объяснимости. Важно заранее согласовать индексацию цен и «выходную» стратегию.'
    ))

    add_heading(document, '8. Инструменты и практики')
    add_paragraph(document, (
        'E‑procurement (Ariba, Coupa и аналоги), FinOps (теги, бюджеты, chargeback, резервации), '
        'SBOM (SPDX/CycloneDX), управление зависимостями и уязвимостями (SCA), зрелость цепочки поставок (SLSA), '
        'каталоги стандартных решений для быстрого и безопасного выбора.'
    ))

    add_heading(document, '9. Нормативные аспекты (кратко)')
    add_paragraph(document, (
        'ЕС: EU AI Act (вступление с 2025), NIS2, GDPR, DORA (для финсектора). США: NIST AI RMF, SP 800‑53, SSDF (SP 800‑218), SOC 2. '
        'Россия (для госсектора): 44‑ФЗ и 223‑ФЗ, импортозамещение, требования к облакам и ПО. '
        'Корпоративные стандарты: ISO/IEC 27001:2022, ITIL 4, COBIT.'
    ))

    add_heading(document, 'Заключение')
    add_paragraph(document, (
        'Грамотная организация IT‑закупок — это баланс скорости и контроля. Лучший подход — гибридная модель: '
        'центр компетенций задаёт стандарты и проводит сложные сделки, а команды оперативно закупают типовые решения '
        'по «белым спискам». Особое внимание — ИИ‑сервисам (данные, риски, объяснимость) и стратегии выхода '
        '(экспорт данных и стоимость миграции). Цель — устойчивость, безопасность и предсказуемость затрат.'
    ))

    add_heading(document, 'Список источников')
    sources = [
        'ISO/IEC 27001:2022 — официальный сайт ISO: https://www.iso.org/standard/27001',
        'SOC 2 Trust Services Criteria — AICPA: https://www.aicpa.org/interestareas/frc/assuranceadvisoryservices/socforserviceorganizations.html',
        'NIST AI Risk Management Framework (AI RMF 1.0): https://www.nist.gov/ai/rmf',
        'NIST SP 800‑53 Rev. 5: https://csrc.nist.gov/publications/detail/sp/800-53/rev-5/final',
        'NIST SP 800‑218 (SSDF): https://csrc.nist.gov/publications/detail/sp/800-218/final',
        'EU AI Act — Еврокомиссия: https://digital-strategy.ec.europa.eu',
        'Директива NIS2 — Еврокомиссия: https://digital-strategy.ec.europa.eu/en/policies/nis2-directive',
        'DORA — ЕС: https://finance.ec.europa.eu/regulation-and-supervision/financial-services-legislation/digital-operational-resilience-act-dora_en',
        'GDPR — Европейская комиссия: https://commission.europa.eu/law/law-topic/data-protection_en',
        'FinOps Foundation: https://www.finops.org/',
        'SPDX: https://spdx.dev/',
        'CycloneDX: https://cyclonedx.org/',
        'SLSA: https://slsa.dev/',
        'CISA — SBOM: https://www.cisa.gov/sbom',
        'ITIL 4 — AXELOS: https://www.axelos.com/best-practice-solutions/itil',
        'COBIT — ISACA: https://www.isaca.org/resources/cobit',
        '44‑ФЗ и 223‑ФЗ — pravo.gov.ru'
    ]
    for s in sources:
        add_paragraph(document, f'- {s}')

    return document


if __name__ == '__main__':
    doc = build_document()
    output_path = 'Реферат_Организация_закупок_в_IT_2025.docx'
    doc.save(output_path)
    print(f'Created: {output_path}')
