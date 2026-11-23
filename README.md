# RED SMS · VBA для Excel

Выберите пример для своей версии настольного Excel. Все запросы используют подпись `sig`.

## Excel для Mac

Понадобятся [RedSms.bas](macos/RedSms.bas) и [RedSms.applescript](macos/RedSms.applescript).

1. В Терминале из корня репозитория скомпилируйте AppleScript в папку Excel.

   ```sh
   script_dir="$HOME/Library/Application Scripts/com.microsoft.Excel"
   mkdir -p "$script_dir"
   /usr/bin/osacompile -o "$script_dir/RedSms.scpt" macos/RedSms.applescript
   ```
2. Откройте редактор VBA в Excel и импортируйте `RedSms.bas` через **File → Import File**. Если хотите сохранить макросы в книге, используйте формат `.xlsm`.
3. Запустите `ConfigureRedSms` из списка макросов и введите логин и API-ключ. Пример хранит их в памяти Excel.
4. Откройте **View → Immediate Window**, чтобы увидеть HTTP-код и тело ответа.

| Метод | Макрос |
|---|---|
| `GET /api/client/info` | `ClientInfo` проверяет доступ к аккаунту |
| `POST /api/message` | `SendSms` просит имя отправителя, номер и подтверждение отправки одного платного SMS |
| `GET /api/message/{uuid}` | `MessageStatus` просит UUID из ответа на отправку |

При отправке HTTP 200 означает, что сообщение создано. Доставку подтверждает статус `delivered`. При HTTP 420 проверьте баланс. После тайм-аута при отправке исход запроса может быть неизвестен, и повторный POST может создать дубликат.

## Excel для Windows

Вариант Windows пока проходит отдельную проверку. [Черновик Windows](https://github.com/redsms/api-samples-vba/pull/1)

[Документация RED SMS HTTP API](https://docs.redsms.ru/http/api/)
