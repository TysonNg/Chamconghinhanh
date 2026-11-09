"""Stable report ordering by verified employee ID, independently of worker completion."""


def sort_report_persons(persons, employee_order):
    ranks = {employee_id: index for index, employee_id in enumerate(employee_order or [])}
    return sorted(persons, key=lambda person: (
        ranks.get(person.get('employee_id'), len(ranks)),
        str(person.get('name') or '').casefold(), str(person.get('id') or ''),
    ))


def filter_paused_persons(persons, paused_employee_ids):
    paused = set(paused_employee_ids or [])
    return [person for person in persons if person.get('employee_id') not in paused]
