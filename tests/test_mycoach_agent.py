"""MyCoach 常駐程式：長輪詢、每 10 分鐘確認登入、自動安裝 launchd。執行：venv/bin/python -m unittest discover -s tests"""
import json
import plistlib
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import mycoach_client as mc  # noqa: E402
import report_worker as rw  # noqa: E402


class KeepaliveTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "s.json"
        self.path.write_text(json.dumps({
            "004": {"cookies": {"JSESSIONID": "a"}, "loggedIn": True},
            "068": {"cookies": {"JSESSIONID": "b"}, "loggedIn": True},
            "027": {"cookies": {"JSESSIONID": "c"}, "loggedIn": True},
            "003": {"cookies": {"JSESSIONID": "d"}, "loggedIn": False},
        }))
        self.calls = []
        patches = [mock.patch.object(mc, "SESSIONS_PATH", self.path),
                   mock.patch.object(rw, "worker", lambda cfg, m, p, body=None, timeout=30: self.calls.append((p, body)) or {}),
                   mock.patch.object(rw, "log", lambda m: None)]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)

    def tearDown(self):
        self.tmp.cleanup()

    def test_alive_dead_unreachable(self):
        def alive(s):
            sid = s.cookies.get("JSESSIONID")
            if sid == "c":
                raise ConnectionError("timeout")
            return sid == "a"
        with mock.patch.object(mc, "session_alive", alive):
            rw.mycoach_keepalive({})
        self.assertEqual(sorted(self.calls, key=lambda c: c[1]["shop_id"]),
                         [("/api/mac/mycoach/alive", {"shop_id": "004", "alive": True}),
                          ("/api/mac/mycoach/alive", {"shop_id": "068", "alive": False})])   # 連不上的 027 不回報
        data = json.loads(self.path.read_text())
        self.assertEqual({k: v["loggedIn"] for k, v in data.items()}, {"004": True, "068": False, "027": True, "003": False})


class LoopTest(unittest.TestCase):
    def test_long_poll_and_exit_on_update(self):
        stamps = iter([1, 1, 1, 2])
        seen = []
        with mock.patch.object(rw, "_code_stamp", lambda: next(stamps)), \
                mock.patch.object(rw, "mycoach_keepalive", lambda cfg: seen.append("alive")), \
                mock.patch.object(rw, "process_mycoach", lambda cfg, wait=0: seen.append(wait) or True), \
                mock.patch.object(rw, "log", lambda m: None):
            rw.mycoach_loop({})
        self.assertEqual(seen, ["alive", rw.MYCOACH_WAIT_SEC, rw.MYCOACH_WAIT_SEC])   # 10 分鐘內只確認一次

    def test_wait_param(self):
        got = {}
        def fake(cfg, m, p, body=None, timeout=30):
            got.update(path=p, timeout=timeout)
            return {"item": None}
        with mock.patch.object(rw, "worker", fake):
            self.assertFalse(rw.process_mycoach({}, wait=25))
        self.assertEqual(got, {"path": "/api/mac/mycoach/next?wait=25", "timeout": 55})


class AgentTest(unittest.TestCase):
    def test_install_once(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertTrue(rw.ensure_mycoach_agent({}, agents_dir=d, load=False))
            self.assertFalse(rw.ensure_mycoach_agent({}, agents_dir=d, load=False))
            with open(Path(d) / f"{rw.MYCOACH_AGENT_LABEL}.plist", "rb") as f:
                p = plistlib.load(f)
            self.assertEqual(p["ProgramArguments"][1:], ["report_worker.py", "--mycoach"])
            self.assertTrue(p["KeepAlive"])
        with tempfile.TemporaryDirectory() as d:
            self.assertFalse(rw.ensure_mycoach_agent({"mycoachAgent": False}, agents_dir=d, load=False))


if __name__ == "__main__":
    unittest.main()
