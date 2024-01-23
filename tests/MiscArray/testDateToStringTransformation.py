from ..util import TestCaseWithFunctionBook


class TestDateToStringTransformation(TestCaseWithFunctionBook):
    def test_1(self) -> None:
        import datetime as dt

        def time_helper(date: dt.datetime) -> dt.datetime:
            return date + dt.timedelta(hours=2)

        func_DateToStringTransformation = self.book.macro(
            "MiscArray.DateToStringTransformation"
        )

        self.assertEqual(
            ((100.2, "2021-01-02"), (2.1, "2021-01-28")),
            func_DateToStringTransformation(
                [
                    [100.2, (dt.datetime(2021, 1, 2, 2, 0, 0))],
                    [2.1, (dt.datetime(2021, 1, 28, 2, 0, 0))],
                ]
            ),
        )

        self.assertEqual(
            (1.2, 2.1, "2021-03-28"),
            func_DateToStringTransformation(
                [1.2, 2.1, dt.datetime(2021, 3, 28, 10, 2, 10)]
            ),
        )

        self.assertEqual(
            "2021-01",
            func_DateToStringTransformation(
                [dt.datetime(2021, 1, 28, 10, 2, 10)], "yyyy-mm"
            )[0],
        )

        self.assertEqual(
            "2021/01/28",
            func_DateToStringTransformation(
                [dt.datetime(2021, 1, 28, 10, 2, 10)], "yyyy/mm/dd"
            )[0],
        )

        self.assertEqual(
            "2021-01-28 10:02:10",
            func_DateToStringTransformation(
                [time_helper(dt.datetime(2021, 1, 28, 10, 2, 10))],
                "yyyy-mm-dd hh:mm:ss",
            )[0],
        )
