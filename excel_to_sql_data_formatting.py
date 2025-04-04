from openpyxl import load_workbook

workbook = load_workbook(r'/Users/engelbertpereira/Downloads/New Norm.xlsx')
sheets = workbook.worksheets
passenger = sheets[1]
ticket = sheets[2]
bus = sheets[3]
route = sheets[4]
operator = sheets[5]
fare = sheets[6]
reservation = sheets[7]


def convert(*attributes):
    rows = []
    for attribute in attributes:
        row = [element.value for element in attribute]
        rows.append(row)

    return rows


def passenger_details(table=passenger):
    id = table["A"]
    email = table["B"]
    first_name = table["C"]
    last_name = table["D"]

    rows = convert(id, email, first_name, last_name)
    print(rows)
    zipped_data = [(a, b, c, d) for a, b, c, d in zip(*rows)]
    for i in zipped_data:
        print(str(i), end=',\n')


def ticket_details(table=ticket):
    id = table["A"]
    booking_date = table["B"]
    booking_time = table["C"]
    dep_date = table["D"]
    arrival_date = table["E"]
    dep_time = table["F"]
    arrival_time = table["G"]
    dep_station_id = table["H"]
    dep_station_name = table["I"]
    arrival_station_id = table["J"]
    arrival_station_name = table["K"]
    route_id = table["L"]
    seat_type = table["M"]
    seat_id = table["N"]

    rows = convert(id, booking_date, booking_time, dep_date, arrival_date, dep_time, arrival_time, dep_station_id, dep_station_name,
                   arrival_station_id, arrival_station_name, route_id, seat_type, seat_id)

    zipped_data = [(a, b, c, d, e, f, g, h, i, j, k, l, m, n) for a, b, c, d, e, f, g, h, i, j, k, l, m, n in zip(*rows)]
    for i in zipped_data:
        print(str(i), end=',\n')


def bus_details(table=bus):
    id = table["A"]
    company = table["B"]
    capacity = table["C"]

    rows = convert(id, company, capacity)
    zipped_data = [(a, b, c) for a, b, c in zip(*rows)]
    for i in zipped_data:
        print(str(i), end=',\n')


def route_details(table=route):
    route_id = table["A"]
    name = table["B"]
    route_type = table["C"]
    operator_id = table["D"]

    rows = convert(route_id, name, route_type, operator_id)
    zipped_data = [(a, b, c, d) for a, b, c, d in zip(*rows)]
    for i in zipped_data:
        print(str(i), end=',\n')


def operator_details(table=operator):
    id = table["A"]
    name = table["B"]
    contact = table["C"]

    rows = convert(id, name, contact)
    zipped_data = [(a, b, c) for a, b, c in zip(*rows)]
    for i in zipped_data:
        print(str(i), end=',\n')


def fare_details(table=fare):
    id = table["A"]
    type = table["B"]
    fare = table["C"]

    rows = convert(id, type, fare)
    zipped_data = [(a, b, c) for a, b, c in zip(*rows)]
    for i in zipped_data:
        print(str(i), end=',\n')


def reservation_details(table=reservation):
    reservation_id = table["A"]
    passenger_id = table["B"]
    ticket_id = table["C"]
    bus_id = table["D"]
    operator_id = table["E"]

    rows = convert(reservation_id, passenger_id, ticket_id, bus_id, operator_id)
    zipped_data = [(a, b, c, d, e) for a, b, c, d, e in zip(*rows)]
    for i in zipped_data:
        print(str(i), end=',\n')


reservation_details()

