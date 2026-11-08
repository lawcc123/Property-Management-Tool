"""Identity validation from reservation-specific labeled fields (never calendar headers)."""
import re


class BookingIdentityError(ValueError):
    pass


def reservation_identity(property_values, unit_value):
    """Require an explicit supported corporation/property and a valid unit field."""
    corporations = set()
    for value in property_values:
        text = " ".join(str(value).split())
        codes = re.findall(r"\bDEMO\s*[#-]?\s*(\d+)\b", text, re.I)
        if codes:
            if any(code not in {'2002', '1001'} for code in codes):
                raise BookingIdentityError('Reservation details name an unsupported corporation.')
            corporations.update(codes)
        # Only inspect values of property/corporation fields; a shared calendar's
        # navigation heading or description must never supply a resident's identity.
        if re.search(r'\b(?:Example Townhomes|Example Townhomes)\b', text, re.I):
            corporations.add('2002')
        if re.search(r'\b(?:Example Condominium|Example Condominium)\b', text, re.I):
            corporations.add('1001')
    if len(corporations) != 1:
        raise BookingIdentityError('The reservation details do not confirm one resident corporation. Review this booking in BuildingLink.')
    corporation = corporations.pop()
    unit = ' '.join(str(unit_value).strip().split())
    match = re.fullmatch(r'(?:(TH)\s*[-#]?\s*)?(\d+[A-Za-z]?|PH\s*\d+|Retail\s*[A-Za-z0-9]+)', unit, re.I)
    if not match:
        raise BookingIdentityError('The reservation details do not contain a valid unit identifier.')
    if match[1] and corporation != '2002':
        raise BookingIdentityError('The reservation unit and corporation conflict (townhouse versus DEMO 1001).')
    identifier = match[2].upper()
    if corporation == '2002':
        if not re.fullmatch(r'\d+[A-Z]?', identifier):
            raise BookingIdentityError('The DEMO 2002 unit identifier is not a townhouse unit.')
        identifier = 'TH ' + identifier
    return 'DEMO ' + corporation, identifier


def identify_requested_for(detail_text, context_property, listed_requested_for):
    """Read only Requested For, excluding the amenity's shared-from/with notes."""
    match = re.search(r'Requested\s+For\s*:\s*(.*?)\s*Submitted\s+By\s*:', detail_text, re.I | re.S)
    if not match:
        raise BookingIdentityError('Reservation details are missing the Requested For field.')
    requested = ' '.join(match[1].split())
    source = re.search(r'\(\s*from\s+([^()]+)\)\s*$', requested, re.I)
    if re.search(r'\(\s*from\b', requested, re.I) and source is None:
        raise BookingIdentityError('The reservation has an unreadable cross-property resident note.')
    property_values = [source[1]] if source else [context_property]
    plain = requested[:source.start()].strip() if source else requested
    unit_and_name = re.fullmatch(r'(?:Unit\s+)?(.+?)\s+-\s+(.+)', plain, re.I)
    if not unit_and_name:
        raise BookingIdentityError('Reservation Requested For does not identify a unit and resident.')
    corporation, unit = reservation_identity(property_values, unit_and_name[1])
    listed = re.match(r'^(?:Unit\s+)?((?:TH\s*)?\d+[A-Za-z]?|PH\s*\d+|Retail\s*[A-Za-z0-9]+)\s*(?:[\u2022\u00b7]|\s-\s)', listed_requested_for, re.I)
    def unit_key(value):
        return re.sub(r'\s+', '', re.sub(r'^TH\s*', '', value, flags=re.I)).casefold()
    if listed is None or unit_key(listed[1]) != unit_key(unit):
        raise BookingIdentityError('Reservation details do not match the unit in the calendar row.')
    return corporation, unit, f'{unit} - {unit_and_name[2].strip()} ({corporation})'
