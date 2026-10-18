# Лабораторная работа №5: Автоматизация обработки запросов заказчика

**Дисциплина**: Внедрение и сопровождение программного продукта  
**Студент**: Суханов Савелий Сергеевич (гр. 5140904/60401)  
**Преподаватель**: Шишкина В.И.  
**GitHub**: [@RelaxedTuna5](https://github.com/RelaxedTuna5)

---

## 📌 Цель работы
Получить знания и практический опыт применения инструментов автоматизации для первичной обработки клиентских задач с использованием утилиты с открытым исходным кодом [iResolver](https://github.com/kovaloid/iresolver).

---

## 🛠 Стек технологий и окружение
- **Язык / Платформа**: Java 11 (Adoptium OpenJDK 11.0.25 LTS)
- **NLP / Векторизация**: Deeplearning4j 0.9.1, DataVec, алгоритмы Doc2Vec (PV-DBOW + Skip-Gram)
- **Экспертная система**: Apache Drools 7.7.0.Final (диалект MVEL)
- **Интеграция**: Jira REST Java Client (Apache Kafka Jira), Confluence REST API

---

## 📁 Структура репозитория
- `Суханов_СС_Отчет_5.docx` — итоговый отчет по лабораторной работе (ГОСТ / стандарт СПбПУ).
- `iresolver_work/rules/CustomRules.drl` — новые правила Drools (контроль подзадач, перевод в работу, эскалация неназначенных задач).
- `iresolver_work/config/configuration.yml` — конфигурация запросов JQL, векторизатора и процессоров.
- `iresolver_work/output/report.html` — сгенерированный интерактивный HTML-отчет анализа инцидентов.
- `iresolver_work/output/report.txt` — текстовая сводка результатов анализа.
- `iresolver_work/data/` — датасеты и обученные векторные модели Doc2Vec (Issues, Documentation, Confluence).

---

## 🚀 Запуск анализа
```powershell
# Установка переменной окружения Java 11
$env:JAVA_HOME = "C:\jdk11"
$env:PATH = "$env:JAVA_HOME\bin;$env:PATH"

# Запуск iResolver
cd C:\iresolver\bin
.\run.bat
```
