import io
from pathlib import Path

from flask import Response

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle,
)
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
)


def safe_pdf_colour(value, fallback):
    try:
        return colors.HexColor(value or fallback)
    except Exception:
        return colors.HexColor(fallback)


class PDFReport:
    def __init__(
        self,
        *,
        setting,
        title,
        filename,
        orientation='portrait',
        logo_upload_folder=None,
        default_logo_path=None,
        author=None,
        left_margin=14,
        right_margin=14,
        top_margin=12,
        bottom_margin=17,
    ):
        self.setting = setting
        self.title = title
        self.filename = filename
        self.orientation = orientation
        self.logo_upload_folder = (
            Path(logo_upload_folder)
            if logo_upload_folder
            else None
        )
        self.default_logo_path = (
            Path(default_logo_path)
            if default_logo_path
            else None
        )

        self.primary_colour = safe_pdf_colour(
            getattr(setting, 'primary_color', None),
            '#0D6EFD',
        )

        self.secondary_colour = safe_pdf_colour(
            getattr(setting, 'secondary_color', None),
            '#198754',
        )

        self.table_header_colour = safe_pdf_colour(
            getattr(setting, 'table_header_color', None),
            '#EAF2F8',
        )

        self.success_colour = safe_pdf_colour(
            getattr(setting, 'success_color', None),
            '#198754',
        )

        self.warning_colour = safe_pdf_colour(
            getattr(setting, 'warning_color', None),
            '#FFC107',
        )

        self.danger_colour = safe_pdf_colour(
            getattr(setting, 'danger_color', None),
            '#DC3545',
        )

        self.page_size = (
            landscape(A4)
            if orientation == 'landscape'
            else A4
        )

        self.buffer = io.BytesIO()

        self.document = SimpleDocTemplate(
            self.buffer,
            pagesize=self.page_size,
            leftMargin=left_margin * mm,
            rightMargin=right_margin * mm,
            topMargin=top_margin * mm,
            bottomMargin=bottom_margin * mm,
            title=title,
            author=(
                author
                or getattr(
                    setting,
                    'organisation_name',
                    'SL Village Banking Pro',
                )
            ),
        )

        self.styles = getSampleStyleSheet()
        self.elements = []

        self.title_style = ParagraphStyle(
            'PDFReportTitle',
            parent=self.styles['Heading1'],
            fontSize=15,
            leading=18,
            textColor=self.primary_colour,
            spaceBefore=3,
            spaceAfter=8,
        )

        self.section_style = ParagraphStyle(
            'PDFReportSection',
            parent=self.styles['Heading2'],
            fontSize=10,
            leading=13,
            textColor=self.secondary_colour,
            spaceBefore=7,
            spaceAfter=6,
        )

        self.small_style = ParagraphStyle(
            'PDFReportSmall',
            parent=self.styles['Normal'],
            fontSize=7,
            leading=9,
        )

        self.normal_style = ParagraphStyle(
            'PDFReportNormal',
            parent=self.styles['Normal'],
            fontSize=8,
            leading=10,
        )

    def _logo_path(self):
        filename = (
            getattr(self.setting, 'report_logo', None)
            or getattr(self.setting, 'logo', None)
        )

        if filename and self.logo_upload_folder:
            candidate = (
                self.logo_upload_folder
                / Path(filename).name
            )

            if candidate.exists():
                return candidate

        if (
            self.default_logo_path
            and self.default_logo_path.exists()
        ):
            return self.default_logo_path

        return None

    def add_branding(self):
        organisation_name = getattr(
            self.setting,
            'organisation_name',
            'Your Organisation Name',
        )

        motto = getattr(
            self.setting,
            'motto',
            '',
        ) or ''

        registration_number = getattr(
            self.setting,
            'registration_number',
            '',
        ) or ''

        address = getattr(
            self.setting,
            'organization_address',
            '',
        ) or ''

        phone = getattr(
            self.setting,
            'organization_phone',
            '',
        ) or ''

        email = getattr(
            self.setting,
            'organization_email',
            '',
        ) or ''

        name_style = ParagraphStyle(
            'PDFBrandName',
            parent=self.styles['Title'],
            fontSize=16,
            leading=19,
            textColor=self.primary_colour,
            spaceAfter=2,
        )

        motto_style = ParagraphStyle(
            'PDFBrandMotto',
            parent=self.styles['Normal'],
            fontSize=8,
            leading=10,
            textColor=self.secondary_colour,
            italic=True,
        )

        contact_style = ParagraphStyle(
            'PDFBrandContact',
            parent=self.styles['Normal'],
            fontSize=7.5,
            leading=9.5,
            textColor=colors.HexColor('#555555'),
        )

        text_parts = [
            Paragraph(organisation_name, name_style)
        ]

        if motto:
            text_parts.append(
                Paragraph(motto, motto_style)
            )

        contacts = []

        if registration_number:
            contacts.append(
                f'Registration No: {registration_number}'
            )

        if address:
            contacts.append(address)

        if phone:
            contacts.append(phone)

        if email:
            contacts.append(email)

        if contacts:
            text_parts.append(
                Paragraph(
                    ' | '.join(contacts),
                    contact_style,
                )
            )

        logo_cell = ''

        logo_path = self._logo_path()

        if logo_path:
            try:
                logo_cell = Image(
                    str(logo_path),
                    width=24 * mm,
                    height=24 * mm,
                    kind='proportional',
                )
            except Exception:
                logo_cell = ''

        header = Table(
            [[logo_cell, text_parts]],
            colWidths=[
                30 * mm,
                self.document.width - 30 * mm,
            ],
        )

        header.setStyle(
            TableStyle([
                (
                    'VALIGN',
                    (0, 0),
                    (-1, -1),
                    'MIDDLE',
                ),
                (
                    'LEFTPADDING',
                    (0, 0),
                    (-1, -1),
                    0,
                ),
                (
                    'RIGHTPADDING',
                    (0, 0),
                    (-1, -1),
                    5,
                ),
                (
                    'TOPPADDING',
                    (0, 0),
                    (-1, -1),
                    0,
                ),
                (
                    'BOTTOMPADDING',
                    (0, 0),
                    (-1, -1),
                    5,
                ),
                (
                    'LINEBELOW',
                    (0, 0),
                    (-1, -1),
                    1.2,
                    self.primary_colour,
                ),
            ])
        )

        self.elements.append(header)
        self.elements.append(Spacer(1, 7))

        return self

    def add_title(self, text=None):
        self.elements.append(
            Paragraph(
                text or self.title.upper(),
                self.title_style,
            )
        )

        return self

    def add_section(self, text):
        self.elements.append(
            Paragraph(
                text,
                self.section_style,
            )
        )

        return self

    def add_paragraph(self, text, style=None):
        self.elements.append(
            Paragraph(
                text,
                style or self.normal_style,
            )
        )

        return self

    def add_spacer(self, height=6):
        self.elements.append(
            Spacer(1, height)
        )

        return self

    def add_information_table(
        self,
        rows,
        *,
        col_widths=None,
        label_columns=(0, 2),
        font_size=8,
    ):
        table = Table(
            rows,
            colWidths=col_widths,
            hAlign='LEFT',
        )

        style = [
            (
                'FONTSIZE',
                (0, 0),
                (-1, -1),
                font_size,
            ),
            (
                'GRID',
                (0, 0),
                (-1, -1),
                0.4,
                self.primary_colour,
            ),
            (
                'VALIGN',
                (0, 0),
                (-1, -1),
                'MIDDLE',
            ),
            (
                'TOPPADDING',
                (0, 0),
                (-1, -1),
                5,
            ),
            (
                'BOTTOMPADDING',
                (0, 0),
                (-1, -1),
                5,
            ),
        ]

        for column in label_columns:
            style.extend([
                (
                    'BACKGROUND',
                    (column, 0),
                    (column, -1),
                    self.table_header_colour,
                ),
                (
                    'FONTNAME',
                    (column, 0),
                    (column, -1),
                    'Helvetica-Bold',
                ),
            ])

        table.setStyle(
            TableStyle(style)
        )

        self.elements.append(table)

        return self

    def add_data_table(
        self,
        rows,
        *,
        col_widths=None,
        numeric_columns=(),
        repeat_rows=1,
        font_size=7,
        header_font_size=7,
        total_row=False,
    ):
        table = Table(
            rows,
            repeatRows=repeat_rows,
            colWidths=col_widths,
            hAlign='LEFT',
        )

        style = [
            (
                'BACKGROUND',
                (0, 0),
                (-1, 0),
                self.primary_colour,
            ),
            (
                'TEXTCOLOR',
                (0, 0),
                (-1, 0),
                colors.white,
            ),
            (
                'FONTNAME',
                (0, 0),
                (-1, 0),
                'Helvetica-Bold',
            ),
            (
                'FONTSIZE',
                (0, 0),
                (-1, 0),
                header_font_size,
            ),
            (
                'FONTSIZE',
                (0, 1),
                (-1, -1),
                font_size,
            ),
            (
                'VALIGN',
                (0, 0),
                (-1, -1),
                'MIDDLE',
            ),
            (
                'GRID',
                (0, 0),
                (-1, -1),
                0.35,
                colors.HexColor('#AAB7C4'),
            ),
            (
                'ROWBACKGROUNDS',
                (0, 1),
                (-1, -1 if not total_row else -2),
                [
                    colors.white,
                    self.table_header_colour,
                ],
            ),
            (
                'TOPPADDING',
                (0, 0),
                (-1, -1),
                4,
            ),
            (
                'BOTTOMPADDING',
                (0, 0),
                (-1, -1),
                4,
            ),
        ]

        for column in numeric_columns:
            style.append(
                (
                    'ALIGN',
                    (column, 1),
                    (column, -1),
                    'RIGHT',
                )
            )

        if total_row:
            style.extend([
                (
                    'BACKGROUND',
                    (0, -1),
                    (-1, -1),
                    self.secondary_colour,
                ),
                (
                    'TEXTCOLOR',
                    (0, -1),
                    (-1, -1),
                    colors.white,
                ),
                (
                    'FONTNAME',
                    (0, -1),
                    (-1, -1),
                    'Helvetica-Bold',
                ),
            ])

        table.setStyle(
            TableStyle(style)
        )

        self.elements.append(table)

        return table

    def add_signatures(
        self,
        labels,
        *,
        include_dates=False,
    ):
        count = len(labels)

        rows = [
            ['________________________'] * count,
            labels,
        ]

        if include_dates:
            rows.append(
                ['Date: __________________'] * count
            )

        table = Table(
            rows,
            colWidths=[
                self.document.width / count
            ] * count,
        )

        table.setStyle(
            TableStyle([
                (
                    'ALIGN',
                    (0, 0),
                    (-1, -1),
                    'CENTER',
                ),
                (
                    'FONTSIZE',
                    (0, 0),
                    (-1, -1),
                    8,
                ),
                (
                    'TOPPADDING',
                    (0, 0),
                    (-1, -1),
                    3,
                ),
                (
                    'BOTTOMPADDING',
                    (0, 0),
                    (-1, -1),
                    3,
                ),
            ])
        )

        self.elements.append(table)

        return self

    def _draw_footer(self, canvas, document):
        canvas.saveState()

        organisation_name = getattr(
            self.setting,
            'organisation_name',
            'Your Organisation Name',
        )

        product_name = getattr(
            self.setting,
            'product_name',
            'SL Village Banking Pro',
        )

        developer_name = getattr(
            self.setting,
            'developer_name',
            'SL Consulting Limited',
        )

        page_width, _ = document.pagesize

        canvas.setStrokeColor(
            self.primary_colour
        )

        canvas.setLineWidth(0.6)

        canvas.line(
            document.leftMargin,
            13 * mm,
            page_width - document.rightMargin,
            13 * mm,
        )

        canvas.setFont(
            'Helvetica',
            7,
        )

        canvas.setFillColor(
            colors.HexColor('#555555')
        )

        canvas.drawString(
            document.leftMargin,
            8 * mm,
            (
                f'{organisation_name} | '
                f'{product_name}'
            ),
        )

        canvas.drawRightString(
            page_width - document.rightMargin,
            8 * mm,
            (
                f'Page {canvas.getPageNumber()} | '
                f'Produced by {developer_name}'
            ),
        )

        canvas.restoreState()

    def build(self):
        self.document.build(
            self.elements,
            onFirstPage=self._draw_footer,
            onLaterPages=self._draw_footer,
        )

        self.buffer.seek(0)

        return self.buffer

    def response(
        self,
        *,
        inline=True,
    ):
        self.build()

        disposition = (
            'inline'
            if inline
            else 'attachment'
        )

        return Response(
            self.buffer.getvalue(),
            mimetype='application/pdf',
            headers={
                'Content-Disposition':
                    (
                        f'{disposition}; '
                        f'filename="{self.filename}"'
                    )
            },
        )