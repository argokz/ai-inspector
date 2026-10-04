#!/usr/bin/env python3
"""
MedReportValidator - Система валидации студенческих отчетов КМУ
с поддержкой OpenXML (.docx) и смыслового анализа Gemini AI.
"""

import os
import re
import sys
import json
import zipfile
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional

def load_env() -> Dict[str, str]:
    """Считывает переменные из .env файла без сторонних библиотек."""
    env_vars = {}
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, val = line.split("=", 1)
                    env_vars[key.strip()] = val.strip().strip("'\"")
    return env_vars

@dataclass
class DocxLayoutInfo:
    is_docx: bool = False
    has_page_borders: bool = False
    borders_first_page_only: bool = False
    has_title_page_setting: bool = False
    left_margin_cm: float = 0.0
    right_margin_cm: float = 0.0
    top_margin_cm: float = 0.0
    bottom_margin_cm: float = 0.0

def extract_from_file(file_path: str) -> Tuple[str, DocxLayoutInfo]:
    """Извлекает текст, структуру таблиц и параметры верстки из .docx или .txt."""
    layout = DocxLayoutInfo()
    if file_path.lower().endswith(".docx"):
        layout.is_docx = True
        try:
            with zipfile.ZipFile(file_path) as docx_zip:
                xml_content = docx_zip.read("word/document.xml")
            tree = ET.fromstring(xml_content)
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            
            body = tree.find(".//w:body", ns)
            lines = []
            if body is not None:
                for elem in body:
                    tag = elem.tag.split("}")[-1]
                    if tag == "p":
                        t = "".join([node.text for node in elem.iterfind(".//w:t", ns) if node.text])
                        if t.strip():
                            lines.append(t)
                    elif tag == "tbl":
                        for row in elem.iterfind(".//w:tr", ns):
                            cells = []
                            for cell in row.iterfind(".//w:tc", ns):
                                c_text = " ".join([node.text for node in cell.iterfind(".//w:t", ns) if node.text])
                                cells.append(c_text.strip())
                            lines.append("| " + " | ".join(cells) + " |")

            full_text = "\n".join(lines)

            # Проверка рамок страницы (w:pgBorders)
            pg_borders = tree.findall(".//w:pgBorders", ns)
            if pg_borders:
                layout.has_page_borders = True
                for pb in pg_borders:
                    display_attr = pb.attrib.get(f"{{{ns['w']}}}display", "")
                    if display_attr == "firstPage":
                        layout.borders_first_page_only = True

            # Проверка колонтитула (w:titlePg)
            if tree.findall(".//w:titlePg", ns):
                layout.has_title_page_setting = True

            # Проверка полей страницы (w:pgMar)
            pg_mar = tree.find(".//w:pgMar", ns)
            if pg_mar is not None:
                left_dxa = float(pg_mar.attrib.get(f"{{{ns['w']}}}left", 0))
                right_dxa = float(pg_mar.attrib.get(f"{{{ns['w']}}}right", 0))
                top_dxa = float(pg_mar.attrib.get(f"{{{ns['w']}}}top", 0))
                bottom_dxa = float(pg_mar.attrib.get(f"{{{ns['w']}}}bottom", 0))
                
                layout.left_margin_cm = round(left_dxa / 567.0, 1)
                layout.right_margin_cm = round(right_dxa / 567.0, 1)
                layout.top_margin_cm = round(top_dxa / 567.0, 1)
                layout.bottom_margin_cm = round(bottom_dxa / 567.0, 1)

            return full_text, layout
        except Exception as e:
            print(f"Ошибка при чтении .docx файла: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(), layout

@dataclass
class CriterionResult:
    name: str
    max_score: float
    score: float = 0.0
    passed: bool = False
    comments: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    penalties: List[str] = field(default_factory=list)

class ReportValidator:
    def __init__(self, text: str, layout: DocxLayoutInfo = None):
        self.raw_text = text
        self.text_lower = text.lower()
        self.layout = layout or DocxLayoutInfo()
        self.results: Dict[str, CriterionResult] = {}

    def validate_title_page(self) -> CriterionResult:
        res = CriterionResult(name="Титульный лист и метаданные", max_score=10.0)
        score = 0.0
        
        checks = [
            ("Медицинский университет", 1.5, "Вуз указан"),
            ("Кафедра", 1.5, "Кафедра указана"),
            ("Цифровая грамотность", 1.0, "Дисциплина указана"),
            ("Выполнил", 1.5, "Сведения о студенте указаны"),
            ("Проверил", 1.5, "Преподаватель указан"),
            (r"(Караганда|Алматы|Астана|202\d)", 1.0, "Город и актуальный год указаны")
        ]
        
        for pattern, weight, label in checks:
            if re.search(pattern, self.raw_text, re.IGNORECASE):
                score += weight
                res.comments.append(f"✓ {label}")
            else:
                res.warnings.append(f"✗ Не обнаружено: {label}")

        if self.layout.is_docx:
            if self.layout.has_page_borders:
                if self.layout.borders_first_page_only:
                    score += 2.0
                    res.comments.append("✓ Рамка титульного листа установлена (только для 1-й страницы)")
                else:
                    score += 1.0
                    res.warnings.append("! Граница страницы включена, но не ограничена первой страницей")
            else:
                res.penalties.append("✗ Нарушение п. 6.1.4: рамка титульного листа (Границы страниц) отсутствует!")

            if self.layout.has_title_page_setting:
                res.comments.append("✓ Включен 'Особый колонтитул для первой страницы' (номер скрыт на титульнике)")
            else:
                res.warnings.append("! Не включен 'Особый колонтитул для первой страницы' (номер может отображаться на титульнике)")
        else:
            score += 2.0

        res.score = min(score, res.max_score)
        res.passed = res.score >= 7.0
        return res

    def validate_objectives(self) -> CriterionResult:
        res = CriterionResult(name="Цель и задачи исследования", max_score=5.0)
        has_goal = bool(re.search(r"цель\s*:", self.text_lower))
        has_tasks = bool(re.search(r"задачи\s*:", self.text_lower))
        
        tasks_count = len(re.findall(r"\n\s*\d+\.\s+[^\n]+", self.raw_text[:3500]))
        
        if has_goal:
            res.score += 2.0
            res.comments.append("✓ Цель исследования сформулирована")
        else:
            res.warnings.append("✗ Цель исследования не найдена")

        if has_tasks and tasks_count >= 5:
            res.score += 3.0
            res.comments.append(f"✓ Задачи декомпозированы (выявлено пунктов: {tasks_count})")
        elif has_tasks:
            res.score += 1.5
            res.warnings.append(f"! Слишком мало задач ({tasks_count} из рекомендуемых >= 5)")
        else:
            res.warnings.append("✗ Задачи исследования не сформулированы")

        res.passed = res.score >= 3.5
        return res

    def validate_medical_directions(self) -> CriterionResult:
        res = CriterionResult(name="Аннотации статей по 5 направлениям", max_score=20.0)
        directions = [
            ("Эпидемиология", r"эпидемиологи"),
            ("Этиология / Патогенез", r"этиологи"),
            ("Диагностика", r"диагностик"),
            ("Профилактика", r"профилактик"),
            ("Терапия / Лечение", r"(терапи|лечени)")
        ]
        
        found_dirs = 0
        for dir_name, pattern in directions:
            if re.search(pattern, self.text_lower):
                found_dirs += 1
                res.comments.append(f"✓ Направление найдено: {dir_name}")
            else:
                res.warnings.append(f"✗ Отсутствует раздел: {dir_name}")
                
        citations = re.findall(r"\[\d+\]", self.raw_text)
        unique_citations = set(citations)
        
        dir_score = (found_dirs / len(directions)) * 15.0
        res.score += dir_score
        
        if len(unique_citations) >= 5:
            res.score += 5.0
            res.comments.append(f"✓ Наличие перекрёстных ссылок [X] на статьи: {len(unique_citations)} шт.")
        else:
            res.warnings.append(f"! Недостаточно ссылок на литературу в обзоре ({len(unique_citations)} из мин. 5)")
            res.score += (len(unique_citations) / 5.0) * 5.0

        res.score = min(res.score, res.max_score)
        res.passed = found_dirs == 5 and len(unique_citations) >= 5
        return res

    def validate_statistics_tables(self) -> CriterionResult:
        res = CriterionResult(name="Статистические таблицы и данные", max_score=20.0)
        score = 0.0
        
        t1 = bool(re.search(r"таблица\s*1", self.text_lower))
        t2 = bool(re.search(r"таблица\s*2", self.text_lower))
        t3 = bool(re.search(r"таблица\s*3", self.text_lower))
        
        if t1:
            score += 4.0
            res.comments.append("✓ Таблица 1 (Демография) присутствует")
        else:
            res.warnings.append("✗ Таблица 1 отсутствует")
            
        if t2:
            score += 4.0
            res.comments.append("✓ Таблица 2 (Заболеваемость/смертность) присутствует")
        else:
            res.warnings.append("✗ Таблица 2 отсутствует")

        if t3:
            score += 4.0
            res.comments.append("✓ Таблица 3 (Региональное распределение) присутствует")
            
            t3_match = re.search(r"таблица\s*3.*?(таблица|работа в researchrabbit|рис|\Z)", self.raw_text, re.IGNORECASE | re.DOTALL)
            if t3_match:
                table_lines = [line for line in t3_match.group(0).split('\n') if '|' in line]
                data_rows = max(0, len(table_lines) - 1)
                if data_rows >= 15:
                    score += 4.0
                    res.comments.append(f"✓ В Таблице 3 соблюдено требование >= 15 строк (найдено: {data_rows})")
                elif data_rows > 0:
                    score += (data_rows / 15.0) * 4.0
                    res.penalties.append(f"! В Таблице 3 меньше 15 строк ({data_rows} строк). Требуется не менее 15 городов/регионов!")
                else:
                    # Если таблицы оформлены без pipe-разметки
                    score += 2.0
                    res.warnings.append("! Не удалось автоматически подсчитать строки Таблицы 3")
        else:
            res.warnings.append("✗ Таблица 3 отсутствует")

        icd_found = bool(re.search(r"(мкб-11|мкб\s*11|icd-11|icd\s*11)", self.text_lower))
        if icd_found:
            score += 4.0
            res.comments.append("✓ Официальный код заболевания по МКБ-11 указан")
        else:
            res.penalties.append("✗ КРИТИЧНО: Отсутствует код заболевания по МКБ-11 (ICD-11) в описании набора данных (штраф -4 балла)")

        if "википедия" in self.text_lower or "wikipedia" in self.text_lower:
            penalty = 3.0
            score = max(0.0, score - penalty)
            res.penalties.append(f"! Нарушение академической чистоты: Википедия использована как источник официальной статистики (-{penalty} балла)")

        res.score = min(score, res.max_score)
        res.passed = (t1 and t2 and t3 and icd_found)
        return res

    def validate_researchrabbit(self) -> CriterionResult:
        res = CriterionResult(name="Раздел ResearchRabbit", max_score=10.0)
        score = 0.0
        
        has_section = bool(re.search(r"researchrabbit", self.text_lower))
        has_collection = bool(re.search(r"(коллекци|библиотек|добавлен)", self.text_lower))
        has_graph = bool(re.search(r"(граф|сет[ьи]|связ|карт[аы])", self.text_lower))
        has_figures = len(re.findall(r"(рис\.|рисунок)\s*\d+", self.raw_text, re.IGNORECASE))
        
        if has_section:
            score += 3.0
            res.comments.append("✓ Подраздел ResearchRabbit присутствует")
        else:
            res.warnings.append("✗ Подраздел ResearchRabbit не найден")
            
        if has_collection:
            score += 2.0
            res.comments.append("✓ Описан процесс создания коллекции/добавления статей")
            
        if has_graph:
            score += 2.5
            res.comments.append("✓ Описан анализ графа связей / Similar Work")
            
        if has_figures >= 3:
            score += 2.5
            res.comments.append(f"✓ Приведены скриншоты этапов с нумерацией (найдено рисунков: {has_figures})")
        else:
            res.warnings.append("! Недостаточно скриншотов этапов в ResearchRabbit (рекомендуется от 3)")
            score += 1.0

        res.score = min(score, res.max_score)
        res.passed = score >= 7.0
        return res

    def validate_ai_usage(self) -> CriterionResult:
        res = CriterionResult(name="Использование ИИ и промт-инжиниринг", max_score=15.0)
        score = 0.0
        
        has_ai_sec = bool(re.search(r"использовани[ея]\s+искусственного\s+интеллекта", self.text_lower))
        has_model = bool(re.search(r"(chatgpt|gpt-4|gemini|claude|copilot)", self.text_lower))
        
        prompt_match = re.search(r"промт\s*:\s*[«\"]([^»\"]+)[»\"]", self.raw_text, re.IGNORECASE)
        prompt_text = prompt_match.group(1) if prompt_match else ""
        has_verification = bool(re.search(r"(провер|сверк|сопостав|верификац)", self.text_lower))

        if has_ai_sec:
            score += 2.0
            res.comments.append("✓ Специализированный подраздел по ИИ оформлен")
        else:
            res.warnings.append("✗ Раздел по ИИ отсутствует")

        if has_model:
            score += 2.0
            res.comments.append("✓ Модель/версия ИИ зафиксирована")
        else:
            res.warnings.append("! Не указана конкретная модель ИИ")

        if prompt_text:
            score += 3.0
            res.comments.append(f"✓ Текст промта задокументирован: \"{prompt_text[:60]}...\"")
            
            components = 0
            if re.search(r"ты\s*-\s*|специалист|эксперт|роль", prompt_text, re.IGNORECASE): components += 1
            if re.search(r"найди|исследуй|составь|проанализируй", prompt_text, re.IGNORECASE): components += 1
            if re.search(r"за последние|австрали|казахстан|период|год", prompt_text, re.IGNORECASE): components += 1
            if re.search(r"в виде|таблиц|списк|формат", prompt_text, re.IGNORECASE): components += 1
            if re.search(r"не используй|только|огранич|провер", prompt_text, re.IGNORECASE): components += 1

            if components >= 4:
                score += 3.0
                res.comments.append("✓ Промт отвечает академической 5-компонентной структуре")
            else:
                score += 1.5
                res.warnings.append(f"! Промт упрощенный (выявлено компонентов структуры: {components}/5). Рекомендуется добавлять явный Формат и Ограничения.")
        else:
            res.warnings.append("✗ Полный текст ключевого промта не приведен")

        if has_verification:
            score += 3.0
            res.comments.append("✓ Описана процедура проверки и сопоставления с первоисточниками")
        else:
            res.warnings.append("✗ Не описан фактчекинг ответов ИИ по базам данных")

        if not re.search(r"(дата использования|запрос выполнен|обращение к ии).*?202\d", self.text_lower):
            res.warnings.append("! Не указана точная дата обращения к ИИ (требование п. 4.5)")
        else:
            score += 2.0

        res.score = min(score, res.max_score)
        res.passed = score >= 10.0
        return res

    def validate_conclusions(self) -> CriterionResult:
        res = CriterionResult(name="Выводы отчёта", max_score=10.0)
        score = 0.0
        
        has_concl = bool(re.search(r"выводы\b", self.text_lower))
        if not has_concl:
            res.warnings.append("✗ Раздел 'Выводы' не найден")
            return res

        score += 3.0
        res.comments.append("✓ Раздел 'Выводы' присутствует")
        
        has_stats_concl = bool(re.search(r"(динамик|заболеваемост|смертност|пик|статистик)", self.text_lower))
        has_limits = bool(re.search(r"(ограничени|галлюцинац|недостоверн|ошибк|mesh|doi)", self.text_lower))
        has_synthesis = bool(re.search(r"(эпидемиолог|вакцинац|профилактик|терапи)", self.text_lower))
        
        if has_stats_concl:
            score += 2.5
            res.comments.append("✓ Содержит анализ медико-статистических трендов")
        if has_limits:
            score += 2.5
            res.comments.append("✓ Отражены ограничения ИИ-инструментов и баз данных")
        if has_synthesis:
            score += 2.0
            res.comments.append("✓ Сформулирован синтез по клинико-эпидемиологическим направлениям")

        res.score = min(score, res.max_score)
        res.passed = score >= 7.0
        return res

    def validate_bibliography(self) -> CriterionResult:
        res = CriterionResult(name="Библиографический список (AMA Style)", max_score=10.0)
        score = 0.0
        
        has_bib = bool(re.search(r"(библиографический список|список литературы)", self.text_lower))
        if not has_bib:
            res.warnings.append("✗ Список литературы не найден")
            return res

        score += 3.0
        res.comments.append("✓ Библиографический список сформирован")
        
        bib_part = self.raw_text[self.raw_text.lower().rfind("библиографический список"):]
        lines = [l.strip() for l in bib_part.split('\n') if len(l.strip()) > 10]
        
        ama_entries = [l for l in lines if re.search(r"doi:\s*10\.|published\s*202\d|\d{4};\d+", l, re.IGNORECASE)]
        bare_urls = [l for l in lines if re.match(r"^https?://[^\s]+$", l.strip())]

        if len(ama_entries) >= 5:
            score += 5.0
            res.comments.append(f"✓ Научные статьи корректно оформлены по стилю AMA (найдено: {len(ama_entries)})")
        else:
            score += 2.0
            res.warnings.append(f"! Мало источников в строгом стиле AMA ({len(ama_entries)})")

        if bare_urls:
            penalty = min(3.0, len(bare_urls) * 0.75)
            score = max(0.0, score - penalty)
            res.penalties.append(f"✗ Нарушение оформления: обнаружено {len(bare_urls)} 'голых' ссылок URL без библиографического описания (-{penalty:.1f} б.)")
        else:
            score += 2.0
            res.comments.append("✓ 'Голые' ссылки отсутствуют, все интернет-ресурсы имеют полное библиографическое описание")

        res.score = min(score, res.max_score)
        res.passed = score >= 6.5
        return res

    def evaluate_with_gemini(self, api_key: str) -> Optional[Dict]:
        """Отправляет ключевые части отчета в Gemini для смыслового рецензирования."""
        if not api_key:
            return None

        prompt = f"""
Ты — строгий ассоциированный профессор и рецензент кафедры информатики и биостатистики медицинского университета.
Оцени студенческий отчет по дисциплине «Цифровая грамотность и использование ИИ».

ТЕКСТ ОТЧЕТА СТУДЕНТА:
\"\"\"
{self.raw_text[:12000]}
\"\"\"

ТВОЯ ЗАДАЧА:
1. Проверить смысловую адекватность статей по 5 направлениям (Эпидемиология, Этиология, Диагностика, Профилактика, Терапия).
2. Оценить профессионализм промт-инжиниринга студента (наличие роли, контекста, инструкций, ограничений).
3. Оценить глубину выводов (не «вода» ли).
4. Дать персонализированную рецензию-отзыв преподавателя для студента.

ОТВЕТ ВЕРНИ СТРОГО В ВИДЕ JSON СО СЛЕДУЮЩЕЙ СТРУКТУРОЙ:
{{
  "semantic_quality_score": <число от 0 до 100>,
  "directions_verdict": "<краткий вердикт по 5 направлениям>",
  "prompt_engineering_review": "<оценка промта>",
  "conclusions_depth_review": "<оценка выводов>",
  "teacher_feedback": "<подробный вежливый текст рецензии преподавателя студенту с указанием плюсов и минусов>",
  "key_recommendations": ["<рекомендация 1>", "<рекомендация 2>"]
}}
"""
        models_to_try = ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
        for model in models_to_try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
            payload = json.dumps({
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"responseMimeType": "application/json"}
            }).encode("utf-8")

            req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"})
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    resp_data = json.loads(resp.read().decode("utf-8"))
                    text_resp = resp_data["candidates"][0]["content"]["parts"][0]["text"]
                    return json.loads(text_resp)
            except Exception as e:
                # Пробуем следующую модель
                continue

        return {"error": "Не удалось подключиться к Gemini API (проверьте интернет-соединение или токен в .env)"}

    def run_all(self, gemini_key: str = "") -> Dict:
        self.results["title"] = self.validate_title_page()
        self.results["objectives"] = self.validate_objectives()
        self.results["directions"] = self.validate_medical_directions()
        self.results["statistics"] = self.validate_statistics_tables()
        self.results["researchrabbit"] = self.validate_researchrabbit()
        self.results["ai_usage"] = self.validate_ai_usage()
        self.results["conclusions"] = self.validate_conclusions()
        self.results["bibliography"] = self.validate_bibliography()

        total_score = sum(r.score for r in self.results.values())
        max_total = sum(r.max_score for r in self.results.values())
        percentage = (total_score / max_total) * 100.0

        if percentage >= 95: grade_letter, gpa = "A", 4.0
        elif percentage >= 90: grade_letter, gpa = "A-", 3.67
        elif percentage >= 85: grade_letter, gpa = "B+", 3.33
        elif percentage >= 80: grade_letter, gpa = "B", 3.0
        elif percentage >= 75: grade_letter, gpa = "B-", 2.67
        elif percentage >= 70: grade_letter, gpa = "C+", 2.33
        elif percentage >= 65: grade_letter, gpa = "C", 2.0
        elif percentage >= 60: grade_letter, gpa = "C-", 1.67
        elif percentage >= 55: grade_letter, gpa = "D+", 1.33
        elif percentage >= 50: grade_letter, gpa = "D", 1.0
        else: grade_letter, gpa = "F", 0.0

        gemini_review = None
        if gemini_key:
            gemini_review = self.evaluate_with_gemini(gemini_key)

        return {
            "total_score": round(total_score, 1),
            "max_total": max_total,
            "percentage": round(percentage, 1),
            "grade_letter": grade_letter,
            "gpa": gpa,
            "details": self.results,
            "gemini_review": gemini_review
        }

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="MedReportValidator: Проверка отчетов по регламенту КМУ")
    parser.add_argument("file", nargs="?", help="Путь к файлу отчета (.txt, .md, .docx)")
    parser.add_argument("--json", action="store_true", help="Вывод в формате JSON")
    parser.add_argument("--ai", action="store_true", help="Включить смысловую рецензию Gemini AI")
    args = parser.parse_args()

    env = load_env()
    gemini_key = env.get("GEMINI_API_KEY", "") if args.ai else ""

    if args.file:
        content, layout = extract_from_file(args.file)
    else:
        content = sys.stdin.read()
        layout = DocxLayoutInfo()

    validator = ReportValidator(content, layout)
    report = validator.run_all(gemini_key=gemini_key)

    if args.json:
        json_dict = {
            "total_score": report["total_score"],
            "max_total": report["max_total"],
            "percentage": report["percentage"],
            "grade_letter": report["grade_letter"],
            "gpa": report["gpa"],
            "criteria": {
                k: {
                    "name": v.name,
                    "score": round(v.score, 1),
                    "max_score": v.max_score,
                    "passed": v.passed,
                    "comments": v.comments,
                    "warnings": v.warnings,
                    "penalties": v.penalties
                } for k, v in report["details"].items()
            },
            "gemini_review": report["gemini_review"]
        }
        print(json.dumps(json_dict, ensure_ascii=False, indent=2))
        sys.exit(0)

    print("=" * 70)
    print("ЭКСПЕРТНЫЙ ЛИСТ ВАЛИДАЦИИ ОТЧЁТА (КМУ)")
    print("=" * 70)
    print(f"ИТОГОВЫЙ БАЛЛ: {report['total_score']} / {report['max_total']} ({report['percentage']}%)")
    print(f"ОЦЕНКА: {report['grade_letter']} (GPA: {report['gpa']})")
    print("=" * 70)
    
    for key, crit in report["details"].items():
        status_symbol = "✔" if crit.passed else "✖"
        print(f"\n[{status_symbol}] {crit.name.upper()}: {crit.score:.1f} / {crit.max_score:.1f} б.")
        for comm in crit.comments:
            print(f"   {comm}")
        for warn in crit.warnings:
            print(f"   {warn}")
        for pen in crit.penalties:
            print(f"   {pen}")

    if report.get("gemini_review") and "teacher_feedback" in report["gemini_review"]:
        print("\n" + "=" * 70)
        print("РЕЦЕНЗИЯ ПРЕПОДАВАТЕЛЯ (GEMINI AI):")
        print("=" * 70)
        print(report["gemini_review"]["teacher_feedback"])

    print("\n" + "=" * 70)
    print("РЕКОМЕНДАЦИИ ПО ДОРАБОТКЕ:")
    has_issues = False
    for crit in report["details"].values():
        if crit.penalties or [w for w in crit.warnings if w.startswith("✗")]:
            for p in crit.penalties:
                print(f" • {p}")
                has_issues = True
            for w in crit.warnings:
                if w.startswith("✗"):
                    print(f" • {w}")
                    has_issues = True
    if not has_issues:
        print(" • Отчёт полностью удовлетворяет методическим требованиям!")
    print("=" * 70)
