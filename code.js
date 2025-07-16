function forMargin() {
  const apiKey = "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjUwMTIwdjEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTc1NTU0NzczOCwiaWQiOiIwMTk1MTJmNC0wYjY2LTdhYjAtODJjNy0xNzFiMjUyZDA0NTkiLCJpaWQiOjE4MTAzMTc3Mywib2lkIjo0MjA1MTIyLCJzIjo3OTM0LCJzaWQiOiJiNmQxNzZiNS1jNmQzLTQxNzQtOGIxMi01MmQxM2RmZDUxODMiLCJ0IjpmYWxzZSwidWlkIjoxODEwMzE3NzN9.1xW2eJ0LuhN8qonZDPJiWfBGJ57Lbc1fmdHKkbEOMBFhGc1ywYWPOy6w25xC_XjNJwg__biYMQPqhhXvHPV9DA";
  const wb_cards = "https://content-api.wildberries.ru/content/v2/get/cards/list";

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet_margin = spreadsheet.getSheetByName("Маржа");
  let sheet_cost = spreadsheet.getSheetByName("Себестоимость");
  sheet_margin.clear();
  sheet_margin.appendRow(['Артикул', 'Название', 'Себестоимость', 'Продажи', 'Возвраты', 'Логистика', 'Комиссия', 'Штрафы', 'Выручка моя', 'Выручка WB', 'Перечисление', 'Итог', 'Маржа']);

  let glob_logistic = 0; // стоимость логистики за товар
  let glob_comission = 0; // стоимость комиссим за товар
  let glob_revenue_my = 0; // размер моей выручки за товар
  let glob_revenue_wb = 0; // размер выручки wb за товар
  let glob_transfer = 0; // к перечислению продавцу за реализованный товар
  let glob_count_order = 0; // количество заказов за товар
  let glob_count_return = 0; // количество возвратов за товар
  let glob_margin = 0; // посчитанная маржа (по формуле) за товар
  let glob_count_margin = 0;
  let glob_result = 0; // итого к оплате за товар (с вычетом логистики и вообще всего)
  let glob_penalty = 0; // общая сумма штрафов

// ---------------------------------------------------------------------------------------------------------------------------------------------------
// получаем значения себестоимости из второго листа
  let cost_list = [];
  for(let index = 1; index <= sheet_cost.getLastRow(); index++) {
    const cell1 = sheet_cost.getRange("A" + index.toString());
    const cell2 = sheet_cost.getRange("B" + index.toString());
    cost_list.push([cell1.getValue(), cell2.getValue()]);
  }

// ---------------------------------------------------------------------------------------------------------------------------------------------------
// получаем всю статистику по логистике, комиссии, выручке, к перечислению
  let string_id = 0; // номер строки для продолжения подтягивания отчета частями
  let data_stat_full = []; // полная статистику за весь поставленный диапазон дат
  let wb_statistics = "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod?dateFrom=2025-02-01&dateTo=2025-02-16&limit=100000&rrdid=" + string_id;
  let opt_statistic = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    }
  };
  while(true) {
    const resp_statistic = UrlFetchApp.fetch(wb_statistics, opt_statistic);
    const data_statistic = JSON.parse(resp_statistic.getContentText());
    data_statistic.forEach(item => {
      data_stat_full.push(item);
    });
    if (data_statistic.length < 100000) {
      break;
    }
    Utilities.sleep(60000);
  }

// ---------------------------------------------------------------------------------------------------------------------------------------------------
// обработка вывода артикулов, наименований и прочего
  let cards_list = [];
  let total = 100;
  let payload_cards = {
    "settings": {
      "cursor": {
        "limit": 100
      },
      "filter": {
        "withPhoto": -1
      }
    }
  };
  let opt_cards = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload_cards)
  };
  while (total == 100) {
    const resp_cards = UrlFetchApp.fetch(wb_cards, opt_cards);
    const data_cards = JSON.parse(resp_cards.getContentText());
    data_cards.cards.forEach(item => {
      const article = item.vendorCode; // артикул продавца
      const name = item.title; // наименование товара
      let logistic = 0; // стоимость логистики за товар
      let cost_price = 0; // себестоимость данного товара
      let comission = 0; // стоимость комиссим за товар
      let revenue_my = 0; // размер моей выручки за товар
      let revenue_wb = 0; // размер выручки wb за товар
      let transfer = 0; // к перечислению продавцу за реализованный товар
      let count_order = 0; // количество заказов за товар
      let count_return = 0; // количество возвратов за товар
      let margin = 0; // посчитанная маржа (по формуле) за товар
      let result = 0; // итого к оплате за товар (с вычетом логистики и вообще всего)
      let penalty = 0; // общая сумма штрафов за товар
      cost_list.forEach(it => {
        if (it[0] == article) {
          cost_price = it[1];
        }
      });
      data_stat_full.forEach(stat => {
        if (stat.sa_name == article) {
          logistic += stat.delivery_rub;
          comission += stat.ppvz_vw + stat.acquiring_fee + stat.ppvz_vw_nds;
          revenue_my += stat.retail_price;
          revenue_wb += stat.retail_amount;
          transfer += stat.ppvz_for_pay;
          count_order += stat.quantity;
          count_return += stat.return_amount;
          penalty += stat.penalty;
        }
      });
      if (count_order > 0) {
        result = transfer - logistic - cost_price * count_order;
      } else {
        if (count_return > 0) {
          result += -logistic - penalty;
        }
      }
      if (revenue_wb != 0) {
        margin = result / revenue_wb * 100;
      }

      glob_logistic += logistic; // стоимость логистики за товар
      glob_comission += comission; // стоимость комиссим за товар
      glob_revenue_my += revenue_my; // размер моей выручки за товар
      glob_revenue_wb += revenue_wb; // размер выручки wb за товар
      glob_transfer += transfer; // к перечислению продавцу за реализованный товар
      glob_count_order += count_order; // количество заказов за товар
      glob_count_return += count_return; // количество возвратов за товар
      glob_margin += margin; // посчитанная маржа (по формуле) за товар
      glob_count_margin += 1;
      glob_result += result; // итого к оплате за товар (с вычетом логистики и вообще всего)
      glob_penalty += penalty;

      cards_list.push([article, name, cost_price, count_order, count_return, logistic, comission, penalty, revenue_my, revenue_wb, transfer, result, margin]);
    });
    let wbNumber = data_cards.cursor.nmID; // получаем артикул WB с которого начнаем получать следующие сто карточек
    let upd = data_cards.cursor.updatedAt; // получаем дату с которой нужно начинать подтягивать следующие строки
    payload_cards = {
      "settings": {
        "cursor": {
          "updatedAt": upd,
          "nmID": wbNumber,
          "limit": 100
        },
        "filter": {
          "withPhoto": -1
        }
      }
    };
    opt_cards = {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + apiKey
      },
      payload: JSON.stringify(payload_cards)
    };
    total = data_cards.cursor.total;
  }
// ---------------------------------------------------------------------------------------------------------------------------------------------------

  let startCell = sheet_margin.getRange("A2");
  startCell.offset(0, 0, cards_list.length, 13).setValues(cards_list);

  sheet_margin.appendRow(['ИТОГО', '', '', glob_count_order, glob_count_return, glob_logistic, glob_comission, glob_penalty, glob_revenue_my, glob_revenue_wb, glob_transfer, glob_result, glob_margin / glob_count_margin]);

}
