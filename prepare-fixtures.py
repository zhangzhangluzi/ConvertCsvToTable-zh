import csv
import io
import json
from pathlib import Path

root = Path(__file__).resolve().parent / "build" / "fixtures"
root.mkdir(exist_ok=True)
cases = []
for name, encoding in [("utf8_no_bom", "utf-8"), ("utf8_bom", "utf-8-sig"), ("gbk", "gbk"), ("utf16_bom", "utf-16")]:
    suffix = "😀" if encoding != "gbk" else ""
    rows = [
        ["start-date", "timezone", "app-id", "creative-id", "spend"],
        ["2026/11/03", "Asia/Hong_Kong", "测试游戏 Android", '中文素材_300×250_逗号,引号"A"' + suffix, "182.165346"],
        ["2026/11/03", "Asia/Hong_Kong", "测试游戏 iOS", "第二行\n换行测试", "24402.028673"],
        ["2026/11/03", "Asia/Hong_Kong", "Rummy RUSH", "ASCII_material", "0.0"],
    ]
    text = io.StringIO(newline="")
    csv.writer(text, lineterminator="\r\n").writerows(rows)
    path = root / f"{name}.csv"
    path.write_bytes(text.getvalue().encode(encoding))
    cases.append({"name": name, "path": str(path), "rows": rows})
(root / "expected.json").write_text(json.dumps(cases, ensure_ascii=False, indent=2), encoding="utf-8")
print("Prepared four encoding fixtures with Chinese, quotes, comma, newline and decimal amounts.")
