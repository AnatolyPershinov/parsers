# МСГ Парсеры

Всего три группы парсеров, каждая из групп парсит свой формат данных

## 1. [Парсеры ТИП 2015,2017](/Parsers20152017)
Структура json файла
``` json
{
    "file_name": "str",
    "object": "str",
    "object_1": "str",
    "work": {
        "work title": "str",
        "work id": "str",
        "upper works": "str",
        "measurements": "str",
        "amount": "str",
        "work_data": {
            "start_date": {
                "plan": "str",
                "estimate": "str",
                "fact": "str"
            },
            "stop_date": {
                "plan": "str",
                "estimate": "str",
                "fact": "str"
            },
            "complite_state_perc": {
                "plan": "float",
                "fact": "float"
            },
            "complite_state_value": {
                "plan": "float",
                "fact": "float"
            },
            "current_remain_perc": "NoneType",
            "current_remain_value": "int",
            "whole_remain_perc": "float",
            "whole_remain_value": "float",
            "mounth_complite_value": {
                "plan": "float",
                "fact": "float"
            },
            "mounth_complite_perc": {
                "plan": "float",
                "fact": "float"
            },
            "progress": [
                {
                    "01.01.2017": {
                        "plan": "float",
                        "fact": "float"
                    }
                },
                {
                    "02.01.2017": {
                        "plan": "float",
                        "fact": "float"
                    }
                },
                ...
                {
                    "31.01.2017": {
                        "plan": "float",
                        "fact": "float"
                    }
                }
            ],
            "comments": "str"
        }
    },
    "resource": {
        "resource_id": "int",
        "resource_name": "str",
        "type": "str",
        "progress": [
            {
                "1": "float"
            },
            {
                "2": "float"
            },
            ...
            {
                "31": "float"
            }
        ],
        "comments": "NoneType"
    }
}
```

## 2. [Парсеры ТИП 2016,2017](/Parsers20162017)


## 3. [Парсеры \(Мессояха и ННГ\)](/ParsersMesayhaNNG)


## 4. [Парсеры новый порт](/Parsersother)
