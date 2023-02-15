import unittest
from pathlib import Path

if __name__ == "__main__":
    loader = unittest.TestLoader()
    suite = loader.discover(Path(__file__).parent.as_posix())
    runner = unittest.TextTestRunner()
    result = runner.run(suite)
    exit(0 if result.wasSuccessful() else 1)
