KICAD_SCRIPTING_PATH: str = '/usr/share/kicad/plugins'
DOCUMENT_COMPONY_NAME: str = 'Horns and Hoovers'
FONT_NAME: str = 'PT Sans'
FONT_SIZE: int = 12
PRIMARY_COLOR: str = '#74B357'
SECONDARY_COLOR: str = '#CCFDB6'
OPTIONAL_COLOR: str = '#A5ECFF'
EXCHANGEABLE_COLOR: str = '#FFE277'
TOTAL_COLOR: str = '#BE9DFC'
SHOPS: dict = {
    'chipdip': 'https://www.chipdip.ru/search?searchtext={}',
    'promelec': 'https://www.promelec.ru/product/{}/',
    'elitan': 'https://www.elitan.ru/price/item{}',
    'robotshop': 'https://shop.robotclass.ru/index.php?route=product/product&product_id={}',
    'lcsc': 'https://lcsc.com/search?q={}'
}


def make_hyperlink(shop_name: str, param: str):
    if len(param) == 0:
        return '—'

    if shop_name in SHOPS.keys():
        return f'=HYPERLINK("{SHOPS[shop_name].format(param)}", "{param}")'
    return param
