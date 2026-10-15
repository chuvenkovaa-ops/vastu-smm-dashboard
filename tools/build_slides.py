"""Собирает слайды презентации стратегии для вкладки «Стратегия».

Запуск: python3 tools/build_slides.py путь/к/стратегии.pdf
Нужны pdftoppm и pdftotext (пакет poppler-utils).
Результат: папка slides/ с картинками s-001.jpg … и slides/slides.js (число слайдов и разделы).
"""
import glob, json, os, re, subprocess, sys

def main(pdf):
    os.makedirs("slides", exist_ok=True)
    for f in glob.glob("slides/s-*.jpg"):
        os.remove(f)
    subprocess.run(["pdftoppm", "-r", "120", "-jpeg", "-jpegopt", "quality=74", pdf, "slides/s"], check=True)
    files = sorted(glob.glob("slides/s-*.jpg"))
    width = len(str(len(files)))
    for f in files:  # pdftoppm сам выбирает ширину номера; приводим к s-001.jpg
        n = int(re.search(r"-(\d+)\.jpg$", f).group(1))
        os.rename(f, "slides/s-%03d.jpg" % n)
    count = len(files)

    # Разделы: слайд-разделитель начинается с номера «01»…«09» и названия
    sections = []
    for p in range(1, count + 1):
        txt = subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), "-layout", pdf, "-"],
                             capture_output=True, text=True).stdout
        lines = [l.strip() for l in txt.splitlines() if l.strip()]
        if len(lines) >= 2 and re.fullmatch(r"0\d", lines[0]) and len(lines) <= 6:
            title = lines[1]
            if len(lines) > 2 and re.search(r"\s(по|и|на|в)$", title):  # заголовок перенесён на вторую строку
                title += " " + lines[2]
            sections.append([lines[0], title, p])
    data = {"count": count, "sections": sections}
    with open("slides/slides.js", "w", encoding="utf-8") as f:
        f.write("/* Собрано tools/build_slides.py */\nwindow.SLIDES = " + json.dumps(data, ensure_ascii=False) + ";\n")
    print(count, "слайдов;", "разделы:", sections)

if __name__ == "__main__":
    main(sys.argv[1])
