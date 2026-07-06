import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APP_PATH = ROOT / "app.py"


def load_app():
    spec = importlib.util.spec_from_file_location("bond_app", APP_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module.app


app = load_app()


class RecognizeAgentPayloadTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def test_text_field_is_supported(self):
        response = self.client.post(
            "/api/recognize-agent",
            json={"text": "245523.SH 山能YK02 1000 净价100 中信信托信昱13号 i020055109"},
        )
        self.assertEqual(response.status_code, 200)
        body = response.get_json()
        self.assertTrue(body["success"])
        self.assertEqual(body["count"], 1)

    def test_input_text_alias_is_supported(self):
        response = self.client.post(
            "/api/recognize-agent",
            json={"input_text": "245523.SH 山能YK02 1000 净价100 中信信托信昱13号 i020055109"},
        )
        self.assertEqual(response.status_code, 200)
        body = response.get_json()
        self.assertTrue(body["success"])
        self.assertEqual(body["count"], 1)


if __name__ == "__main__":
    unittest.main()
