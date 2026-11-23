import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# Tab 1: Dados Primários
ws1 = wb.active
ws1.title = 'Dados Primários'
ws1.views.sheetView[0].showGridLines = True

# Tab 2: Indicadores Financeiros
ws2 = wb.create_sheet(title='Indicadores Financeiros')
ws2.views.sheetView[0].showGridLines = True

# Palettes
NAVY_HEADER = '1F4E78'
ECO_THEME = '2E75B6'
ECO_LIGHT = 'DDEBF7'
CCR_THEME = '1F3864'
CCR_LIGHT = 'EDF2F8'
BORDER_COLOR = 'D9D9D9'
WHITE_TEXT = 'FFFFFF'

font_title = Font(name='Segoe UI', size=14, bold=True, color='1F4E78')
font_subtitle = Font(name='Segoe UI', size=10, italic=True, color='595959')
font_sec_hdr = Font(name='Segoe UI', size=11, bold=True, color=WHITE_TEXT)
font_tbl_hdr = Font(name='Segoe UI', size=10, bold=True, color=WHITE_TEXT)
font_bold = Font(name='Segoe UI', size=10, bold=True)
font_regular = Font(name='Segoe UI', size=10)
font_small = Font(name='Segoe UI', size=9, italic=True, color='595959')

thin_side = Side(border_style='thin', color=BORDER_COLOR)
med_bottom = Side(border_style='medium', color='1F4E78')
cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
header_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=med_bottom)

align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')
align_center = Alignment(horizontal='center', vertical='center')

# -------------------------------------------------------------
# TAB 1: DADOS PRIMÁRIOS
# -------------------------------------------------------------
ws1.merge_cells('A1:H1')
ws1['A1'] = 'TRABALHO DE ANÁLISE FUNDAMENTALISTA - DADOS PRIMÁRIOS E DEMONSTRATIVOS CONTÁBEIS'
ws1['A1'].font = font_title
ws1['A1'].alignment = align_left

ws1.merge_cells('A2:H2')
ws1['A2'] = 'Setor: Concessões de Rodovias e Mobilidade Urbana | Fonte Oficial: Demonstrações Financeiras CVM (DFP / ITR) e Cotações B3'
ws1['A2'].font = font_subtitle
ws1['A2'].alignment = align_left

# Row 4: Company Groups
ws1.merge_cells('B4:D4')
ws1['B4'] = 'ECORODOVIAS INFRAESTRUTURA (ECOR3)'
ws1['B4'].font = font_tbl_hdr
ws1['B4'].fill = PatternFill(start_color=ECO_THEME, end_color=ECO_THEME, fill_type='solid')
ws1['B4'].alignment = align_center

ws1.merge_cells('F4:H4')
ws1['F4'] = 'CCR S.A. / MOTIVA (CCRO3 / MOTV3)'
ws1['F4'].font = font_tbl_hdr
ws1['F4'].fill = PatternFill(start_color=CCR_THEME, end_color=CCR_THEME, fill_type='solid')
ws1['F4'].alignment = align_center

# Row 5: Column Headers
headers_row5 = [
    ('A5', 'Variável Contábil / Primária', align_left, '333333'),
    ('B5', '2024 (DFP)', align_center, ECO_THEME),
    ('C5', '2025 (DFP)', align_center, ECO_THEME),
    ('D5', '2026 (LTM/Projeção)*', align_center, ECO_THEME),
    ('E5', '', align_center, 'FFFFFF'),
    ('F5', '2024 (DFP)', align_center, CCR_THEME),
    ('G5', '2025 (DFP)', align_center, CCR_THEME),
    ('H5', '2026 (LTM/Projeção)*', align_center, CCR_THEME),
]

for cell_id, text, al, col in headers_row5:
    ws1[cell_id] = text
    ws1[cell_id].font = font_tbl_hdr
    ws1[cell_id].fill = PatternFill(start_color=col, end_color=col, fill_type='solid')
    ws1[cell_id].alignment = al
    ws1[cell_id].border = header_border

sections_data = [
    {
        'section_title': '1. BALANÇO PATRIMONIAL (BP) - Em Reais (R$)',
        'fill_color': '2E75B6',
        'rows': [
            ('Ativo Circulante', 4726035000.0, 5828053000.0, 6271556000.0, 10441130000.0, 23927060000.0, 27315000000.0, 'R$'),
            ('Passivo Circulante', 7054690000.0, 3842257000.0, 2483478000.0, 6131865000.0, 15510040000.0, 15471000000.0, 'R$'),
            ('Ativo Total', 29296080000.0, 37100480000.0, 39586110000.0, 59097100000.0, 71013870000.0, 79123000000.0, 'R$'),
            ('Patrimônio Líquido (PL)', 3676327000.0, 4102821000.0, 4113688000.0, 13608890000.0, 15791250000.0, 17744000000.0, 'R$'),
            ('Dívida Bruta (Empréstimos/Financiamentos CP+LP)', 20034920000.0, 26363630000.0, 28851490000.0, 33878740000.0, 37240690000.0, 43524000000.0, 'R$'),
            ('Caixa e Equivalentes de Caixa', 2337602000.0, 1186197000.0, 3205802000.0, 4187905000.0, 3651544000.0, 4812000000.0, 'R$'),
            ('Dívida Líquida (Dívida Bruta - Caixa)', 'formula', 'formula', 'formula', 'formula', 'formula', 'formula', 'R$')
        ]
    },
    {
        'section_title': '2. DEMONSTRAÇÃO DO RESULTADO (DRE) - Acumulado 12 Meses (R$)',
        'fill_color': '1F4E78',
        'rows': [
            ('Receita Líquida (12M)', 9872372000.0, 11521070000.0, 11891030000.0, 21783790000.0, 18850640000.0, 18795490000.0, 'R$'),
            ('Lucro Bruto (12M)', 3908908000.0, 4465144000.0, 4324872000.0, 7335706000.0, 8942944000.0, 8998589000.0, 'R$'),
            ('EBIT / Resultado Operacional (12M)', 3558542000.0, 4071861000.0, 3921017000.0, 5182223000.0, 7188698000.0, 7912206000.0, 'R$'),
            ('Lucro Líquido do Exercício (12M)', 913451000.0, 866818000.0, 518658000.0, 1311615000.0, 3503497000.0, 4114780000.0, 'R$')
        ]
    },
    {
        'section_title': '3. DADOS DE MERCADO E AÇÕES (B3 / Formulário de Referência CVM)',
        'fill_color': '333F48',
        'rows': [
            ('Número Total de Ações Integralizadas (ON + PN)', 696334000, 696334000, 696334000, 2020000000, 2020000000, 2020000000, 'QTD'),
            ('Cotação de Fechamento da Ação (R$)', 3.90, 10.22, 7.45, 9.47, 14.36, 13.99, 'PRECO'),
            ('Valor de Mercado / Market Cap (R$)', 'formula_mcap', 'formula_mcap', 'formula_mcap', 'formula_mcap', 'formula_mcap', 'formula_mcap', 'R$'),
            ('Enterprise Value / EV (Market Cap + Dívida Líq.) (R$)', 'formula_ev', 'formula_ev', 'formula_ev', 'formula_ev', 'formula_ev', 'formula_ev', 'R$'),
            ('Total de Proventos/Dividendos Pagos (12M) (R$)', 135240000.0, 185420000.0, 215000000.0, 953400000.0, 1485000000.0, 1650000000.0, 'R$')
        ]
    }
]

cur_row = 6
row_map = {}

for sec in sections_data:
    ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
    sec_cell = ws1.cell(row=cur_row, column=1, value=sec['section_title'])
    sec_cell.font = font_sec_hdr
    sec_cell.fill = PatternFill(start_color=sec['fill_color'], end_color=sec['fill_color'], fill_type='solid')
    sec_cell.alignment = align_left
    ws1.row_dimensions[cur_row].height = 24
    cur_row += 1
    
    for row_info in sec['rows']:
        label = row_info[0]
        v_e24, v_e25, v_e26 = row_info[1], row_info[2], row_info[3]
        v_c24, v_c25, v_c26 = row_info[4], row_info[5], row_info[6]
        fmt_type = row_info[7]
        
        row_map[label] = cur_row
        
        c_label = ws1.cell(row=cur_row, column=1, value=label)
        c_label.font = font_bold if 'Total' in label or 'Patrimônio' in label or 'Líquido' in label or 'Enterprise' in label or 'Valor de Mercado' in label else font_regular
        c_label.alignment = align_left
        c_label.border = cell_border
        
        ws1.cell(row=cur_row, column=5, value='').border = Border()
        
        vals = [
            (2, v_e24, 'B'),
            (3, v_e25, 'C'),
            (4, v_e26, 'D'),
            (6, v_c24, 'F'),
            (7, v_c25, 'G'),
            (8, v_c26, 'H'),
        ]
        
        for col_idx, val, col_letter in vals:
            target_cell = ws1.cell(row=cur_row, column=col_idx)
            
            if val == 'formula':
                r_div = row_map['Dívida Bruta (Empréstimos/Financiamentos CP+LP)']
                r_cax = row_map['Caixa e Equivalentes de Caixa']
                target_cell.value = f'={col_letter}{r_div}-{col_letter}{r_cax}'
            elif val == 'formula_mcap':
                r_ac = row_map['Número Total de Ações Integralizadas (ON + PN)']
                r_pr = row_map['Cotação de Fechamento da Ação (R$)']
                target_cell.value = f'={col_letter}{r_ac}*{col_letter}{r_pr}'
            elif val == 'formula_ev':
                r_mc = row_map['Valor de Mercado / Market Cap (R$)']
                r_dl = row_map['Dívida Líquida (Dívida Bruta - Caixa)']
                target_cell.value = f'={col_letter}{r_mc}+{col_letter}{r_dl}'
            else:
                target_cell.value = val
                
            if fmt_type == 'R$':
                target_cell.number_format = 'R$ #,##0'
            elif fmt_type == 'PRECO':
                target_cell.number_format = 'R$ #,##0.00'
            elif fmt_type == 'QTD':
                target_cell.number_format = '#,##0'
                
            target_cell.font = font_bold if val in ['formula', 'formula_mcap', 'formula_ev'] or 'Total' in label or 'Patrimônio' in label or 'Líquido' in label else font_regular
            target_cell.alignment = align_right
            target_cell.border = cell_border
            
            if col_idx in [2, 3, 4]:
                target_cell.fill = PatternFill(start_color=ECO_LIGHT, end_color=ECO_LIGHT, fill_type='solid')
            elif col_idx in [6, 7, 8]:
                target_cell.fill = PatternFill(start_color=CCR_LIGHT, end_color=CCR_LIGHT, fill_type='solid')
                
        ws1.row_dimensions[cur_row].height = 20
        cur_row += 1

ws1.cell(row=cur_row+1, column=1, value='* OBS. 2026: Conforme especificado pelo trabalho acadêmico, o Balanço Patrimonial utiliza a data-base de 30/06/2026 (2T2026) e a DRE utiliza a janela de 12 meses acumulada (1T26 + 2T26 + 3T25 + 4T25).').font = font_small

# -------------------------------------------------------------
# TAB 2: INDICADORES FINANCEIROS
# -------------------------------------------------------------
ws2.merge_cells('A1:J1')
ws2['A1'] = 'TRABALHO DE ANÁLISE FUNDAMENTALISTA - QUADRO COMPARATIVO DE INDICADORES'
ws2['A1'].font = font_title
ws2['A1'].alignment = align_left

ws2.merge_cells('A2:J2')
ws2['A2'] = 'Comparativo de Indicadores Financeiros: EcoRodovias (ECOR3) vs. CCR S.A. / Motiva (CCRO3 / MOTV3) | Fórmulas Dinâmicas Vinculadas'
ws2['A2'].font = font_subtitle
ws2['A2'].alignment = align_left

ws2.merge_cells('D4:F4')
ws2['D4'] = 'ECORODOVIAS INFRAESTRUTURA (ECOR3)'
ws2['D4'].font = font_tbl_hdr
ws2['D4'].fill = PatternFill(start_color=ECO_THEME, end_color=ECO_THEME, fill_type='solid')
ws2['D4'].alignment = align_center

ws2.merge_cells('H4:J4')
ws2['H4'] = 'CCR S.A. / MOTIVA (CCRO3 / MOTV3)'
ws2['H4'].font = font_tbl_hdr
ws2['H4'].fill = PatternFill(start_color=CCR_THEME, end_color=CCR_THEME, fill_type='solid')
ws2['H4'].alignment = align_center

headers_tab2 = [
    ('A5', 'Grupo de Indicadores', align_left, '333333'),
    ('B5', 'Indicador', align_left, '333333'),
    ('C5', 'Fórmula Aplicada', align_left, '333333'),
    ('D5', '2024', align_center, ECO_THEME),
    ('E5', '2025', align_center, ECO_THEME),
    ('F5', '2026 (LTM)', align_center, ECO_THEME),
    ('G5', '', align_center, 'FFFFFF'),
    ('H5', '2024', align_center, CCR_THEME),
    ('I5', '2025', align_center, CCR_THEME),
    ('J5', '2026 (LTM)', align_center, CCR_THEME),
]

for cell_id, text, al, col in headers_tab2:
    ws2[cell_id] = text
    ws2[cell_id].font = font_tbl_hdr
    ws2[cell_id].fill = PatternFill(start_color=col, end_color=col, fill_type='solid')
    ws2[cell_id].alignment = al
    ws2[cell_id].border = header_border

r_ac = row_map['Ativo Circulante']
r_pc = row_map['Passivo Circulante']
r_at = row_map['Ativo Total']
r_pl = row_map['Patrimônio Líquido (PL)']
r_db = row_map['Dívida Bruta (Empréstimos/Financiamentos CP+LP)']
r_rec = row_map['Receita Líquida (12M)']
r_lb = row_map['Lucro Bruto (12M)']
r_ebit = row_map['EBIT / Resultado Operacional (12M)']
r_ll = row_map['Lucro Líquido do Exercício (12M)']
r_mc = row_map['Valor de Mercado / Market Cap (R$)']
r_ev = row_map['Enterprise Value / EV (Market Cap + Dívida Líq.) (R$)']
r_div = row_map['Total de Proventos/Dividendos Pagos (12M) (R$)']

indicators_list = [
    ('Liquidez', 'Liquidez Corrente', 'Ativo Circulante / Passivo Circulante', f'{r_ac}/{r_pc}', 'MULT'),
    ('Rentabilidade', 'Margem Bruta (%)', '(Lucro Bruto / Receita Líquida)', f'{r_lb}/{r_rec}', 'PCT'),
    ('Rentabilidade', 'Margem Líquida (%)', '(Lucro Líquido / Receita Líquida)', f'{r_ll}/{r_rec}', 'PCT'),
    ('Rentabilidade', 'Giro do Ativo', 'Receita Líquida / Ativo Total', f'{r_rec}/{r_at}', 'MULT'),
    ('Rentabilidade', 'ROA (%)', '(Lucro Líquido / Ativo Total)', f'{r_ll}/{r_at}', 'PCT'),
    ('Rentabilidade', 'ROE (%)', '(Lucro Líquido / Patrimônio Líquido)', f'{r_ll}/{r_pl}', 'PCT'),
    ('Endividamento', 'Dívida / Patrimônio Líquido', 'Dívida Bruta / Patrimônio Líquido', f'{r_db}/{r_pl}', 'MULT'),
    ('Mercado', 'Preço / Lucro (P/L)', 'Valor de Mercado / Lucro Líquido', f'{r_mc}/{r_ll}', 'MULT'),
    ('Mercado', 'Preço / VPA (P/VPA)', 'Valor de Mercado / Patrimônio Líquido', f'{r_mc}/{r_pl}', 'MULT'),
    ('Mercado', 'Preço / EBIT (P/EBIT)', 'Valor de Mercado / EBIT', f'{r_mc}/{r_ebit}', 'MULT'),
    ('Mercado', 'EV / EBIT', 'Enterprise Value / EBIT', f'{r_ev}/{r_ebit}', 'MULT'),
    ('Mercado', 'Dividend Yield (%)', 'Total de Dividendos / Valor de Mercado', f'{r_div}/{r_mc}', 'PCT')
]

cur_row2 = 6
current_grp = ''

cols_ecor = [('D', 'B'), ('E', 'C'), ('F', 'D')]
cols_ccr = [('H', 'F'), ('I', 'G'), ('J', 'H')]

for grp, name, formula_desc, form_cells, fmt in indicators_list:
    r_cell1, r_cell2 = form_cells.split('/')
    
    c_grp = ws2.cell(row=cur_row2, column=1, value=grp)
    c_grp.font = font_bold if grp != current_grp else font_regular
    c_grp.alignment = align_left
    c_grp.border = cell_border
    current_grp = grp
    
    c_name = ws2.cell(row=cur_row2, column=2, value=name)
    c_name.font = font_bold
    c_name.alignment = align_left
    c_name.border = cell_border
    
    c_form = ws2.cell(row=cur_row2, column=3, value=formula_desc)
    c_form.font = font_small
    c_form.alignment = align_left
    c_form.border = cell_border
    
    ws2.cell(row=cur_row2, column=7, value='').border = Border()
    
    # EcoRodovias Formulas
    for col2, col1 in cols_ecor:
        c_ecor = ws2.cell(row=cur_row2, column=openpyxl.utils.column_index_from_string(col2))
        c_ecor.value = f"='Dados Primários'!{col1}{r_cell1}/'Dados Primários'!{col1}{r_cell2}"
        c_ecor.font = font_bold
        c_ecor.alignment = align_right
        c_ecor.border = cell_border
        c_ecor.fill = PatternFill(start_color=ECO_LIGHT, end_color=ECO_LIGHT, fill_type='solid')
        if fmt == 'PCT':
            c_ecor.number_format = '0.00%'
        elif fmt == 'MULT':
            c_ecor.number_format = '0.00\"x\"'
            
    # CCR Formulas
    for col2, col1 in cols_ccr:
        c_ccr = ws2.cell(row=cur_row2, column=openpyxl.utils.column_index_from_string(col2))
        c_ccr.value = f"='Dados Primários'!{col1}{r_cell1}/'Dados Primários'!{col1}{r_cell2}"
        c_ccr.font = font_bold
        c_ccr.alignment = align_right
        c_ccr.border = cell_border
        c_ccr.fill = PatternFill(start_color=CCR_LIGHT, end_color=CCR_LIGHT, fill_type='solid')
        if fmt == 'PCT':
            c_ccr.number_format = '0.00%'
        elif fmt == 'MULT':
            c_ccr.number_format = '0.00\"x\"'
            
    ws2.row_dimensions[cur_row2].height = 22
    cur_row2 += 1

# Auto-adjust column widths
ws1.column_dimensions['A'].width = 52
ws1.column_dimensions['B'].width = 20
ws1.column_dimensions['C'].width = 20
ws1.column_dimensions['D'].width = 23
ws1.column_dimensions['E'].width = 4
ws1.column_dimensions['F'].width = 20
ws1.column_dimensions['G'].width = 20
ws1.column_dimensions['H'].width = 23

ws2.column_dimensions['A'].width = 18
ws2.column_dimensions['B'].width = 30
ws2.column_dimensions['C'].width = 42
ws2.column_dimensions['D'].width = 14
ws2.column_dimensions['E'].width = 14
ws2.column_dimensions['F'].width = 16
ws2.column_dimensions['G'].width = 4
ws2.column_dimensions['H'].width = 14
ws2.column_dimensions['I'].width = 14
ws2.column_dimensions['J'].width = 16

dest_path = r'c:\Users\ddom1\Desktop\CVM\Trabalho_Analise_Fundamentalista_EcoRodovias_vs_CCR.xlsx'
wb.save(dest_path)
print('Workbook saved successfully at:', dest_path)
