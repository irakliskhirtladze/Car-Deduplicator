from deduplicator import ToyotaDeduplicator


if __name__ == "__main__":
    file = 'SCION.xlsx'
    ignore_columns = ['Дата выпуска', 'Цвет салона', 'Цвет кузова']

    deduplicator = ToyotaDeduplicator(file, ignore_columns)
    deduplicator.deduplicate()
