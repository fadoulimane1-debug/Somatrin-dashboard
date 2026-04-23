from django import template

register = template.Library()


@register.filter
def num_plain(value, decimals=None):
    """Nombre au format français : milliers espace, décimales virgule."""
    if value is None or value == '':
        return ''
    if decimals is None or decimals == '':
        d = 2
    else:
        try:
            d = int(str(decimals))
        except (ValueError, TypeError):
            d = 2
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    fmt = f'{{:,.{max(d, 0)}f}}'
    s = fmt.format(v)
    if d <= 0:
        return s.replace(',', '\u00a0')
    integer, decimal = s.split('.')
    integer = integer.replace(',', '\u00a0')
    return f'{integer},{decimal}'


@register.filter
def matricule_only(value):
    if value is None:
        return ""
    return str(value).split("/", 1)[0].strip()
