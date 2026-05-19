/**
 * Единый файл настроек приложения «Доставка заморозки».
 * Редактируйте только этот файл для смены URL, режимов, правил и начальных справочников.
 *
 * После изменений обновите страницу в браузере (лучше с локальным сервером, не file://).
 */
window.APP_CONFIG = {
  /** Режим при первом открытии: 'horeca' | 'gallery' | 'all' */
  defaultMode: 'horeca',

  /** Режим «Все» — комбинированный экспорт HoReCa + Галереи */
  allMode: {
    label: 'Все',
    exportFilePrefix: 'Комбинированный_заказы'
  },

  /** Ссылка «Создать новое планирование» в Яндекс.Курьер */
  yandexPlanningUrl:
    'https://yandex.ru/courier/companies/31071/depots/all/mvrp/import',

  /**
   * Имена листов в Google Sheets (если у режима нет своих sheetNameByDay).
   * Расписание — по дням недели; справочники — фиксированные имена.
   */
  sheetNames: {
    weekdays: {
      monday: 'понедельник',
      tuesday: 'вторник',
      wednesday: 'среда',
      thursday: 'четверг',
      friday: 'пятница',
      saturday: 'суббота',
      sunday: 'воскресенье'
    },
    vehicles: 'vehicles',
    depots: ['depot', 'depots'],
    startLocations: ['startdata', 'startlocations', 'start_points']
  },

  modes: {
    horeca: {
      label: 'HoReCa',
      /** URL Google Apps Script Web App (JSON по листам). Пока пусто — синхронизация недоступна. */
      sheetEndpoint: 'https://script.google.com/macros/s/AKfycbw2EEKXtoGRPV78d7eum7FtY5HobiaRjb9XEehiVfppm0pQSdM7m6KSAoR-NeMc_wfRyA/exec',
      exportFilePrefix: 'HoReCa_заказы',
      defaultTimeWindow: '10:00-21:00',
      /**
       * Имена листов по дням для этого режима (если отличаются от sheetNames.weekdays).
       * Оставьте null — будут общие weekday-имена.
       */
      sheetNameByDay: null,
      /**
       * Автовыбор стартовых точек — заполните после настройки листа StartData в таблице HoReCa.
       */
      autoSelection: {
        priorityDepotId: '1112',
        priorityStartLocationId: '1111',
        defaultStartLocationId: '1110',
        defaultExtraStartId: '1114'
      },
      /** Пустой режим до первой успешной загрузки из Google Sheets (не читать localStorage). */
      startsEmptyUntilSheetSync: true,
      /**
       * ID водителей, выбранных для экспорта по умолчанию после загрузки из таблицы.
       * Остальные остаются в списке, но без галочки «В экспорт».
       */
      defaultExportVehicleIds: ['424'],
      seed: {
        startLocations: [],
        vehicles: [],
        depots: []
      }
    },

    gallery: {
      label: 'Галереи',
      sheetEndpoint:
        'https://script.google.com/macros/s/AKfycbxZ3vPXULau6UusKKe25-k8YIHQ9Q4g7hGpnlXkuPB9rsYFZfHBugLUO4o7_0cnUhgzYQ/exec',
      exportFilePrefix: 'Заморозка_заказы',
      defaultTimeWindow: '10:00-21:00',
      sheetNameByDay: null,
      autoSelection: {
        priorityDepotId: '1112',
        priorityStartLocationId: '1111',
        defaultStartLocationId: '1110',
        defaultExtraStartId: '1114'
      },
      /**
       * ID водителей для экспорта по умолчанию (колонка ID в листе Vehicles).
       * Остальные видны в списке, но без галочки «В экспорт».
       */
      defaultExportVehicleIds: ['424'],
      seed: {
        startLocations: [],
        vehicles: [],
        depots: []
      }
    }
  }
};
