# MacroInstaller (bundled ODS)

Сборка установщика макросов в один `.ods`:

```bash
cd installer
python3 build_ods_bundled.py
# или с тестами:
python3 build_ods_bundled.py --with-tests
```

Результат: `MacroInstaller_bundled_<версия>_<дата>.ods`.

Документные макросы кнопки меню (`set_menu` + cfg/lib/images blob) прописываются в `Scripts/python/` и в `META-INF/manifest.xml`. Без записи в манифест AlterOffice считает файл повреждённым.


Цепочка

Правки в AO → файлы в профиле, например: ~/.config/alteroffice/5/user/config/aoffice.cfg/modules/acell/images/

lc_imagelist.xml — привязки команд ↔ иконки
Bitmaps/lc_userimages.png — спрайт
Скопировать в репозиторий:

cp ~/.config/alteroffice/5/user/config/aoffice.cfg/modules/acell/images/lc_imagelist.xml macro-lib/set_menu_assets/lc_imagelist.xml
cp ~/.config/alteroffice/5/user/config/aoffice.cfg/modules/acell/images/Bitmaps/lc_userimages.png    macro-lib/set_menu_assets/Bitmaps/lc_userimages.png
Вшить в blob:
python3 macro-lib/bundle_set_menu_images.py
Пересобрать инсталлятор:
cd installer && python3 build_ods_bundled.py
Открывать новый MacroInstaller_bundled_<версия>_….ods — старый ODS не обновится сам.


Пересборка docx:

# зависимости один раз
pip install -r docs/requirements-build.txt
# всё
python3 docs/build_docx.py --all
# или точечно
python3 docs/build_docx.py 05_POSTPROCESS.md 15_PARAM_REFERENCE.md