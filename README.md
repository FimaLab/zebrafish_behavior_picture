# Зебрафиш: визуализация Excel

Streamlit-приложение читает загруженный файл `.xlsx` со структурой как в `картинка.xlsx`, берет данные с `Лист1`, позволяет менять текстовые подписи и значения стрелок, а затем рисует две группы с доброй и злой зебрафиш. Без загрузки файла визуализация не строится.

В приложении есть скачивание единой PNG-картинки с metadata `600 dpi`, а также HTML-версии визуализации. Значения `≈` отображаются как `-`.

## Локальный запуск

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address=0.0.0.0 --server.port=8509
```

Откройте:

```text
http://localhost:8509
```

## Docker

```powershell
docker build -t zebrafish-streamlit .
docker run --rm -p 8509:8509 zebrafish-streamlit
```

После запуска контейнера приложение будет доступно по адресу:

```text
http://localhost:8509
```
