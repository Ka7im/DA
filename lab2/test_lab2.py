# Проверка программы работы 2 на приложенной Excel-матрице.
# Запуск: python -m unittest -v test_lab2.py

# Подключаем средства тестирования, точные дроби и путь к данным.
import unittest
from fractions import Fraction
from pathlib import Path

# Импортируем написанный класс и чтение Excel без запуска демонстрации.
from lab2 import AgglomerativeClustering, read_excel


# Проверяем формулы, последовательности объединений и остановку.
class TestClustering(unittest.TestCase):
    # Один раз считываем исходные расстояния из приложенного файла.
    @classmethod
    def setUpClass(cls):
        cls.matrix, cls.labels = read_excel(
            Path(__file__).with_name("lab2.xlsx"))

    # Сверяем 25 объединений с фиксированным контрольным расчётом.
    def test_known_merge_sequences(self):
        expected = {
            "single": [(3, 8, "8"), (2, 9, "9"), (4, 10, "10"),
                       (7, 11, "10"), (1, 12, "11")],
            "complete": [(3, 8, "8"), (2, 4, "10"), (1, 9, "12"),
                         (6, 10, "12"), (5, 11, "14")],
            "average": [(3, 8, "8"), (2, 4, "10"), (1, 9, "23/2"),
                        (6, 10, "23/2"), (7, 12, "38/3")],
            "method1": [(3, 8, "8"), (2, 4, "10"), (1, 9, "23/2"),
                        (6, 10, "23/2"), (11, 12, "12")],
            "method2": [(3, 8, "8"), (2, 4, "10"), (9, 10, "11"),
                        (1, 11, "12"), (6, 7, "13")],
        }
        for method, control in expected.items():
            with self.subTest(method=method):
                model = AgglomerativeClustering(self.matrix, 3, method)
                model.fit(verbose=False)
                actual = [(s["left"], s["right"], str(s["distance"]))
                          for s in model.history]
                self.assertEqual(actual, control)

    # Проверяем все формулы на девяти значениях с нечётной медианой.
    def test_nine_cross_distances(self):
        # Значения: 15, 12, 15, 15, 12, 13, 9, 10, 14; сумма равна 115.
        expected = {"single": 9, "complete": 15,
                    "average": Fraction(115, 9), "method1": 12, "method2": 13}
        for method, control in expected.items():
            with self.subTest(method=method):
                model = AgglomerativeClustering(self.matrix, 3, method)
                self.assertEqual(model.cluster_distance((0, 2, 7), (1, 3, 5)),
                                 control)

    # Проверяем чётную медиану и сохранение повторяющихся расстояний.
    def test_even_median_and_repetitions(self):
        model = AgglomerativeClustering(self.matrix, 3, "method2")
        # После сортировки 9, 10, 12, 15: медиана = (10 + 12) / 2 = 11.
        self.assertEqual(model.cluster_distance((2, 7), (1, 3)), 11)
        # После сортировки 10, 12, 12, 15: медиана = 12, не медиана множества.
        self.assertEqual(model.cluster_distance((0, 1), (2, 3)), 12)

    # Для кластеров разных размеров учитываем каждую исходную пару.
    def test_average_unequal_cluster_sizes(self):
        model = AgglomerativeClustering(self.matrix, 3, "average")
        self.assertEqual(model.cluster_distance((0, 2, 7), (1,)), 13)

    # Проверяем 40 запусков: каждый метод при n_clusters от 1 до 8.
    def test_all_stopping_counts(self):
        for method in AgglomerativeClustering.METHODS:
            for target in range(1, 9):
                with self.subTest(method=method, target=target):
                    model = AgglomerativeClustering(self.matrix, target, method)
                    result = model.fit(verbose=False)
                    self.assertEqual(len(result), target)
                    self.assertEqual(len(model.history), 8 - target)
                    self.assertEqual(sorted(x for group in result for x in group),
                                     self.labels)
                    for step, state in enumerate(model.history, 1):
                        self.assertEqual(len(state["clusters"]), 8 - step)
                        self.assertEqual(len(state["matrix"]), 8 - step)

    # Убеждаемся, что вход не меняется и повторный запуск воспроизводим.
    def test_repeat_fit_preserves_input(self):
        original = [row[:] for row in self.matrix]
        model = AgglomerativeClustering(self.matrix, 3, "average")
        result = model.fit(verbose=False)
        history = model.history[:]
        self.assertEqual(model.fit(verbose=False), result)
        self.assertEqual(model.history, history)
        self.assertEqual(self.matrix, original)

    # Один объект уже является одним кластером: объединений быть не должно.
    def test_single_observation(self):
        model = AgglomerativeClustering([[0]], 1, labels=["A"])
        self.assertEqual(model.fit(verbose=False), [["A"]])
        self.assertEqual(model.history, [])

    # Нулевые расстояния разных объектов допускаются и не смешиваются с диагональю.
    def test_zero_distance_and_tie_rule(self):
        model = AgglomerativeClustering([[0, 0, 0], [0, 0, 0], [0, 0, 0]], 2)
        model.fit(verbose=False)
        self.assertEqual((model.history[0]["left"], model.history[0]["right"]),
                         (1, 2))

    # Отклоняем неправильные размеры, отрицательные и нечисловые расстояния.
    def test_invalid_matrices(self):
        invalid = [[], [[0, 1]], [[0, 1], [2, 0]], [[1, 2], [2, 0]],
                   [[0, -1], [-1, 0]], [[0, float("nan")], [float("nan"), 0]],
                   [[0, float("inf")], [float("inf"), 0]],
                   [[0, "1"], ["1", 0]], [[0, True], [True, 0]]]
        for matrix in invalid:
            with self.subTest(matrix=matrix):
                with self.assertRaises(ValueError):
                    AgglomerativeClustering(matrix, 1)

    # Отклоняем неверное число кластеров, неизвестный метод и повторные метки.
    def test_invalid_parameters(self):
        for target in [0, 9, 2.5, True]:
            with self.assertRaises(ValueError):
                AgglomerativeClustering(self.matrix, target)
        with self.assertRaises(ValueError):
            AgglomerativeClustering(self.matrix, 3, "unknown")
        with self.assertRaises(ValueError):
            AgglomerativeClustering(self.matrix, 3, labels=["P"] * 8)


# Позволяем запустить проверки как отдельную программу.
if __name__ == "__main__":
    unittest.main(verbosity=2)
