# formatting - header
header_r = {'bold': True, 'font_size': 24, 'align': 'right'}
header_l = {'bold': True, 'font_size': 24, 'align': 'left'}

# formatting - portefeuille
header_color = '#ECA359'
positive_color = '#C6EFCE'
negative_color = '#FFC7CE'
port_header = {'bold': True, 'bg_color': header_color, 'align': 'center'}
port_header_border = {'bold': True, 'bg_color': header_color, 'align': 'center', 'left': 1}
aantal_pos = {'font_color': '#796E63', 'bg_color': positive_color, 'left': 1}
aantal_neg = {'font_color': '#796E63', 'bg_color': negative_color, 'left': 1}
aantal_neutral = {'font_color': '#796E63', 'left': 1}
waarde = {'num_format': '#,##0.00;-#,##0.00;—;@'}
procent_pos = {'num_format': '0%', 'bg_color': positive_color}
procent_neg = {'num_format': '0%', 'bg_color': negative_color}
procent_neutral = {'num_format': '0%'}

# formatting - totals
totaal_font = {'bold': True, 'font_size': 14}
totaal_num = {'bold': True, 'font_size': 14, 'num_format': '€ #,##0.00'}

# formatting - winstverlies
winstverlies_font = {'bold': True, 'font_size': 14, 'bg_color': '#D1CDCC'}
winstverlies_num = {'bold': True, 'font_size': 14, 'num_format': '€ #,##0.00;(€ #,##0.00);0', 'bg_color': '#D1CDCC'}

# formatting - jaaroverzicht
jaaroverzicht_font = {'bg_color': '#D1CDCC'}
jaaroverzicht_num = {'num_format': '€ #,##0.00;(€ #,##0.00);0'}

months_translation = {
    'January': 'Januari',
    'February': 'Februari',
    'March': 'Maart',
    'April': 'April',
    'May': 'Mei',
    'June': 'Juni',
    'July': 'Juli',
    'August': 'Augustus',
    'September': 'September',
    'October': 'Oktober',
    'November': 'November',
    'December': 'December'
}
