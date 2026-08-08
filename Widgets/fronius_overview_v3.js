"use strict";

const { items, rules, time, actions, triggers, cache } = require("openhab");
const LoggerFactory = Java.type("org.slf4j.LoggerFactory");

// Module-scoped rule logger. Initialized at rule execute start.
let RULE_LOGGER = null;
function getLogger() {
  if (RULE_LOGGER) return RULE_LOGGER;
  try {
    return LoggerFactory.getLogger("LH30.pv_fronius_sum_values");
  } catch (e) {
    // Fallback simple logger if LoggerFactory is unavailable
    return {
      trace: (msg) => print(msg),
      debug: (msg) => print(msg),
      info: (msg) => print(msg),
      warn: (msg) => print(msg),
      error: (msg) => print(msg),
    };
  }
}

// Java-Zeitklassen importieren
const JavaZDT = Java.type("java.time.ZonedDateTime");
const ZoneId = Java.type("java.time.ZoneId");
const ChronoField = Java.type("java.time.temporal.ChronoField");
const Instant = Java.type("java.time.Instant");

// Konstanten
const TIMEZONE = "Europe/Berlin";
const MINIMAL_POWER_THRESHOLD = 0.001; // 1 Wh Schwelle
const DEFAULT_ZEITRAUM = "heute";
const VALID_ZEITRAEUME = ["heute", "gestern", "woche", "monat", "jahr"];

// gridDraw:            PV_Meter_EnergyConsumed (kWh, kumuliert)
// gridFeed:            PV_Meter_EnergyProduced (kWh, kumuliert)
// load:                PV_Inverter1_Power (W)
// batteryPower:        PV_Battery_Power (W, < 0 = Laden, > 0 = Entladen)
// production:          PV_Production_Power (W)
// NEW:                 PV_Sum_BatteryCharge (kWh, kumuliert)
//                      PV_Sum_BatteryDischarge (kWh, kumuliert)
const SOURCE_ITEMS = {
  gridDraw: "PV_Meter_EnergyConsumed",
  gridFeed: "PV_Meter_EnergyProduced",
  load: "PV_Load_Power",
  batteryPower: "PV_Battery_Power",
  production: "PV_Production_Power",
  batteryChargeSum: "PV_Sum_BatteryCharge",
  batteryDischargeSum: "PV_Sum_BatteryDischarge",
};

// Logger-Funktionen
function logTrace(message) {
  getLogger().trace(`PV sum values: ${message}`);
}

function logDebug(message) {
  getLogger().debug(`PV sum values: ${message}`);
}

function logInfo(message) {
  getLogger().info(`PV sum values: ${message}`);
}

function logWarn(message) {
  getLogger().warn(`PV sum values: ${message}`);
}

function logError(message) {
  getLogger().error(`PV sum values: ${message}`);
}

function formatEnergyForSummary(value) {
  return value < 1
    ? `${(value * 1000).toFixed(1)}Wh`
    : `${value.toFixed(2)}kWh`;
}
function formatDateTime(timestamp) {
  const date = new Date(timestamp);
  return date.toLocaleString("de-DE", {
    timeZone: TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

function logPvSummary(context) {
  const { zeitraum, startTime, endTime, values } = context;

  logInfo(
    `start=${formatDateTime(startTime.toInstant().toEpochMilli())}, ` +
      `end=${formatDateTime(endTime.toInstant().toEpochMilli())}, ` +
      `erzeugung=${formatEnergyForSummary(values.erzeugung)}, ` +
      `gesamtverbrauch=${formatEnergyForSummary(values.gesamtverbrauch)}, ` +
      `eigenverbrauch=${formatEnergyForSummary(values.eigenverbrauch)}, ` +
      `netzbezug=${formatEnergyForSummary(values.netzbezug)}, ` +
      `netzeinspeisung=${formatEnergyForSummary(values.netzeinspeisung)}`,
  );
}

// Zeit-Funktionen
const getTimeRange = (zeitraum) => {
  const zoneId = ZoneId.of(TIMEZONE);
  const now = JavaZDT.now(zoneId);
  const dataCollectionStart = JavaZDT.of(2026, 1, 1, 0, 0, 0, 0, zoneId);

  const startOfDay = (date) =>
    date.withHour(0).withMinute(0).withSecond(0).withNano(0);
  const endOfDay = (date) =>
    date.withHour(23).withMinute(59).withSecond(59).withNano(999999999);

  let startTime, endTime;

  switch (zeitraum.toLowerCase()) {
    case "heute":
      startTime = startOfDay(now);
      endTime = now;
      break;
    case "gestern":
      startTime = startOfDay(now.minusDays(1));
      endTime = endOfDay(now.minusDays(1));
      break;
    case "woche":
      // Letzte 7 Tage (von vor 6 Tagen bis heute)
      startTime = startOfDay(now.minusDays(6));
      endTime = now;
      break;
    case "monat":
      // Letzte 31 Tage (von vor 30 Tagen bis heute)
      startTime = startOfDay(now.minusDays(30));
      endTime = now;
      break;
    case "jahr":
      startTime = startOfDay(now.minusDays(364));
      endTime = now;
      break;
    default:
      startTime = startOfDay(now);
      endTime = now;
  }

  logDebug(`Ursprünglicher Zeitraum für ${zeitraum}:`);
  logDebug(`  Start: ${formatDateTime(startTime.toInstant().toEpochMilli())}`);
  logDebug(`  Ende:  ${formatDateTime(endTime.toInstant().toEpochMilli())}`);

  if (startTime.isBefore(dataCollectionStart)) {
    startTime = dataCollectionStart;
    logDebug(
      `Startzeit auf Datenerfassungsbeginn angepasst: ${formatDateTime(startTime.toInstant().toEpochMilli())}`,
    );
  }

  logDebug(`Finaler Zeitraum ${zeitraum}:`);
  logDebug(`  Start: ${formatDateTime(startTime.toInstant().toEpochMilli())}`);
  logDebug(`  Ende:  ${formatDateTime(endTime.toInstant().toEpochMilli())}`);

  return [startTime, endTime];
};

function getConfiguredZeitraum() {
  try {
    const zeitraumItem = items.getItem("PV_Zeitraum");

    if (
      !zeitraumItem ||
      zeitraumItem.state === null ||
      zeitraumItem.state === undefined
    ) {
      logWarn(`PV_Zeitraum has no state, using default '${DEFAULT_ZEITRAUM}'`);
      return DEFAULT_ZEITRAUM;
    }

    const rawState = zeitraumItem.state.toString();

    if (!rawState || rawState === "NULL" || rawState === "UNDEF") {
      logWarn(
        `PV_Zeitraum is ${rawState || "empty"}, using default '${DEFAULT_ZEITRAUM}'`,
      );
      return DEFAULT_ZEITRAUM;
    }

    const normalizedState = rawState.trim().toLowerCase();

    if (!VALID_ZEITRAEUME.includes(normalizedState)) {
      logWarn(
        `PV_Zeitraum has unsupported value '${rawState}', using default '${DEFAULT_ZEITRAUM}'`,
      );
      return DEFAULT_ZEITRAUM;
    }

    return normalizedState;
  } catch (error) {
    logWarn(
      `Could not read PV_Zeitraum, using default '${DEFAULT_ZEITRAUM}': ${error}`,
    );
    return DEFAULT_ZEITRAUM;
  }
}

// Hilfsfunktion zum Aktualisieren der Anzeige-Items
const updateDisplayItems = (values) => {
  try {
    logDebug("Starte Update der Display-Items mit folgenden Werten:");
    Object.entries(values).forEach(([key, value]) => {
      logDebug(`${key}: ${value}`);
    });

    const formatValue = (value) =>
      value < 1 ? `${(value * 1000).toFixed(1)} Wh` : `${value.toFixed(2)} kWh`;
    items.getItem("PV_Erzeugung").postUpdate(formatValue(values.erzeugung));
    items
      .getItem("PV_Eigenverbrauch")
      .postUpdate(formatValue(values.eigenverbrauch));
    items
      .getItem("PV_Direktverbrauch")
      .postUpdate(formatValue(values.direktverbrauch));
    items
      .getItem("PV_Batterieladung")
      .postUpdate(formatValue(values.batterieladung));
    items
      .getItem("PV_Netzeinspeisung")
      .postUpdate(formatValue(values.netzeinspeisung));
    items
      .getItem("PV_Gesamtverbrauch")
      .postUpdate(formatValue(values.gesamtverbrauch));
    items
      .getItem("PV_Eigenversorgung")
      .postUpdate(formatValue(values.eigenversorgung));
    items
      .getItem("PV_EigenversorgungPV")
      .postUpdate(formatValue(values.eigenversorgungPV));
    items
      .getItem("PV_EigenversorgungBatterie")
      .postUpdate(formatValue(values.eigenversorgungBatterie));
    items.getItem("PV_Netzbezug").postUpdate(formatValue(values.netzbezug));

    const prozentWerte = {
      eigenverbrauchProzent: 0,
      direktverbrauchProzent: 0,
      batterieladungProzent: 0,
      netzeinspeisungProzent: 0,
      eigenversorgungProzent: 0,
      eigenversorgungPVProzent: 0,
      eigenversorgungBatterieProzent: 0,
      netzbezugProzent: 0,
    };

    if (values.erzeugung > MINIMAL_POWER_THRESHOLD) {
      prozentWerte.eigenverbrauchProzent = Math.round(
        (values.eigenverbrauch / values.erzeugung) * 100,
      );
      prozentWerte.direktverbrauchProzent = Math.round(
        (values.direktverbrauch / values.erzeugung) * 100,
      );
      prozentWerte.batterieladungProzent = Math.round(
        (values.batterieladung / values.erzeugung) * 100,
      );
      prozentWerte.netzeinspeisungProzent = Math.round(
        (values.netzeinspeisung / values.erzeugung) * 100,
      );
    }

    if (values.gesamtverbrauch > MINIMAL_POWER_THRESHOLD) {
      const effectiveNetzbezug =
        values.netzbezug < MINIMAL_POWER_THRESHOLD ? 0 : values.netzbezug;
      const effectiveBatterie = values.eigenversorgungBatterie;
      const effectivePV = values.eigenversorgungPV;

      let batterieProzent = Math.round(
        (effectiveBatterie / values.gesamtverbrauch) * 100,
      );
      let pvProzent = Math.round((effectivePV / values.gesamtverbrauch) * 100);
      let netzbezugProzent = Math.round(
        (effectiveNetzbezug / values.gesamtverbrauch) * 100,
      );

      const total = batterieProzent + pvProzent + netzbezugProzent;
      if (total > 100) {
        const factor = 100 / total;
        batterieProzent = Math.round(batterieProzent * factor);
        pvProzent = Math.round(pvProzent * factor);
        netzbezugProzent = 100 - batterieProzent - pvProzent;
      }

      if (batterieProzent > 95 && netzbezugProzent < 5) {
        batterieProzent = 100;
        netzbezugProzent = 0;
        pvProzent = 0;
      }

      prozentWerte.eigenversorgungProzent = batterieProzent + pvProzent;
      prozentWerte.eigenversorgungBatterieProzent = batterieProzent;
      prozentWerte.eigenversorgungPVProzent = pvProzent;
      prozentWerte.netzbezugProzent = netzbezugProzent;
    }

    items
      .getItem("PV_EigenverbrauchProzent")
      .postUpdate(prozentWerte.eigenverbrauchProzent);
    items
      .getItem("PV_DirektverbrauchProzent")
      .postUpdate(prozentWerte.direktverbrauchProzent);
    items
      .getItem("PV_BatterieladungProzent")
      .postUpdate(prozentWerte.batterieladungProzent);
    items
      .getItem("PV_NetzeinspeisungProzent")
      .postUpdate(prozentWerte.netzeinspeisungProzent);
    items
      .getItem("PV_EigenversorgungProzent")
      .postUpdate(prozentWerte.eigenversorgungProzent);
    items
      .getItem("PV_EigenversorgungPVProzent")
      .postUpdate(prozentWerte.eigenversorgungPVProzent);
    items
      .getItem("PV_EigenversorgungBatterieProzent")
      .postUpdate(prozentWerte.eigenversorgungBatterieProzent);
    items
      .getItem("PV_NetzbezugProzent")
      .postUpdate(prozentWerte.netzbezugProzent);

    logDebug("Display-Items erfolgreich aktualisiert");
  } catch (error) {
    logError(`Fehler beim Update der Display-Items: ${error}`);
  }
};

// kWh-Berechnung aus Summen-Items (kWh, kumuliert)
const getEnergyDeltaFromSumItem = (itemName, startTime, endTime) => {
  try {
    const it = items.getItem(itemName);
    const delta = actions.PersistenceExtensions.deltaBetween(
      it,
      startTime,
      endTime,
      "influxdb",
    );
    if (!delta || !delta.doubleValue) {
      logDebug(`deltaBetween liefert keinen Wert für ${itemName}`);
      return 0;
    }
    const val = delta.doubleValue(); // bereits kWh
    logDebug(`${itemName} deltaBetween: ${val.toFixed(3)} kWh`);
    return val < 0 ? 0 : val;
  } catch (e) {
    logError(`Fehler bei deltaBetween(${itemName}): ${e}`);
    return 0;
  }
};

// kWh-Berechnung aus Leistungs-Items in W → Wh → kWh (über Durchschnitt * Dauer)
const getEnergyFromPowerAverageW = (itemName, startTime, endTime) => {
  try {
    const it = items.getItem(itemName);
    const avg = actions.PersistenceExtensions.averageBetween(
      it,
      startTime,
      endTime,
      "influxdb",
    );
    if (!avg || !avg.doubleValue) {
      logDebug(`averageBetween liefert keinen Wert für ${itemName}`);
      return 0;
    }

    const avgW = avg.doubleValue(); // W
    const durationHours =
      (endTime.toInstant().toEpochMilli() -
        startTime.toInstant().toEpochMilli()) /
      3600000.0;
    const energyWh = avgW * durationHours;
    const energyKWh = energyWh / 1000.0;
    logDebug(
      `${itemName} averageBetween: ${avgW.toFixed(1)} W über ${durationHours.toFixed(3)} h = ${energyKWh.toFixed(3)} kWh`,
    );

    return energyKWh;
  } catch (e) {
    logError(`Fehler bei averageBetween(${itemName}): ${e}`);
    return 0;
  }
};

function normalizeFlowEnergyKWh(energyKWh, options = {}) {
  const {
    itemName = "unknown",
    positiveDirection = "import",
    negativeDirection = "export",
  } = options;

  if (energyKWh > 0) {
    logDebug(
      `${itemName}: positive energy ${energyKWh.toFixed(3)} kWh interpreted as ${positiveDirection}`,
    );
  } else if (energyKWh < 0) {
    logDebug(
      `${itemName}: negative energy ${energyKWh.toFixed(3)} kWh interpreted as ${negativeDirection}`,
    );
  } else {
    logDebug(`${itemName}: zero energy over selected interval`);
  }

  return {
    positive: Math.max(0, energyKWh),
    negative: Math.max(0, -energyKWh),
    net: energyKWh,
  };
}

// kWh-Berechnung für Batterie-Ladung/-Entladung aus den neuen Summen-Items
const getBatteryChargeEnergyFromSum = (startTime, endTime) =>
  getEnergyDeltaFromSumItem(SOURCE_ITEMS.batteryChargeSum, startTime, endTime);

const getBatteryDischargeEnergyFromSum = (startTime, endTime) =>
  getEnergyDeltaFromSumItem(
    SOURCE_ITEMS.batteryDischargeSum,
    startTime,
    endTime,
  );

const calculateCurrentValues = async (startTime, endTime) => {
  try {
    logDebug(`Berechne Werte von ${startTime} bis ${endTime}`);

    const gridDraw = getEnergyDeltaFromSumItem(
      SOURCE_ITEMS.gridDraw,
      startTime,
      endTime,
    );
    const gridFeed = getEnergyDeltaFromSumItem(
      SOURCE_ITEMS.gridFeed,
      startTime,
      endTime,
    );

    const production = getEnergyFromPowerAverageW(
      SOURCE_ITEMS.production,
      startTime,
      endTime,
    );
    const inverterLoadFlow = getEnergyFromPowerAverageW(
      SOURCE_ITEMS.load,
      startTime,
      endTime,
    );

    // NEU: Batterie-Energie aus den Summen-Items (kWh, kumuliert)
    const batteryCharge = getBatteryChargeEnergyFromSum(startTime, endTime);
    const batteryDischarge = getBatteryDischargeEnergyFromSum(
      startTime,
      endTime,
    );

    const loadFlow = normalizeFlowEnergyKWh(inverterLoadFlow, {
      itemName: SOURCE_ITEMS.load,
      positiveDirection: "energy received by inverter",
      negativeDirection: "energy delivered by inverter",
    });

    const load = loadFlow.negative;

    logDebug("Rohe Energiemengen:");
    logDebug(`Netzbezug (gridDraw): ${gridDraw.toFixed(3)} kWh`);
    logDebug(`Netzeinspeisung (gridFeed): ${gridFeed.toFixed(3)} kWh`);
    logDebug(`Produktion (production): ${production.toFixed(3)} kWh`);
    logDebug(
      `Wechselrichter-Lastfluss netto (inverterLoadFlow): ${inverterLoadFlow.toFixed(3)} kWh`,
    );
    logDebug(`Hausverbrauch aus PV_Load_Power (load): ${load.toFixed(3)} kWh`);
    logDebug(`Batterieladung (batteryCharge): ${batteryCharge.toFixed(3)} kWh`);
    logDebug(
      `Batterieentladung (batteryDischarge): ${batteryDischarge.toFixed(3)} kWh`,
    );

    const directConsumptionRaw = production - gridFeed - batteryCharge;
    const directConsumptionFromLoad = load - batteryDischarge - gridDraw;

    logInfo(
      `direct consumption comparison: ` +
        `fromProduction=${directConsumptionRaw.toFixed(3)}kWh, ` +
        `fromLoad=${directConsumptionFromLoad.toFixed(3)}kWh`,
    );

    logInfo(
      `energy balance: production=${production.toFixed(3)}kWh, ` +
        `gridFeed=${gridFeed.toFixed(3)}kWh, ` +
        `batteryCharge=${batteryCharge.toFixed(3)}kWh, ` +
        `directConsumptionRaw=${directConsumptionRaw.toFixed(3)}kWh, ` +
        `load=${load.toFixed(3)}kWh, ` +
        `gridDraw=${gridDraw.toFixed(3)}kWh, ` +
        `batteryDischarge=${batteryDischarge.toFixed(3)}kWh`,
    );
    const values = {
      netzbezug: gridDraw,
      netzeinspeisung: gridFeed,
      batterieladung: batteryCharge,
      eigenversorgungBatterie: batteryDischarge,
      direktverbrauch: Math.max(0, directConsumptionRaw),
      verbrauch: load,
    };

    logDebug("Berechnete Basiswerte:");
    Object.entries(values).forEach(([key, value]) =>
      logDebug(`${key}: ${value.toFixed(3)} kWh`),
    );

    values.eigenverbrauch = values.direktverbrauch + values.batterieladung;
    values.erzeugung = values.netzeinspeisung + values.eigenverbrauch;
    values.eigenversorgungPV = values.direktverbrauch;
    values.eigenversorgung =
      values.eigenversorgungPV + values.eigenversorgungBatterie;
    values.gesamtverbrauch = values.eigenversorgung + values.netzbezug;

    logDebug("Abgeleitete Werte:");
    Object.entries(values).forEach(([key, value]) =>
      logDebug(`${key}: ${value.toFixed(3)} kWh`),
    );

    logDebug(
      `Gesamtverbrauch = Eigenversorgung + Netzbezug: ` +
        `${values.gesamtverbrauch.toFixed(3)} = ` +
        `${values.eigenversorgung.toFixed(3)} + ${values.netzbezug.toFixed(3)}`,
    );
    logDebug(
      `Erzeugung = Netzeinspeisung + Eigenverbrauch: ` +
        `${values.erzeugung.toFixed(3)} = ` +
        `${values.netzeinspeisung.toFixed(3)} + ${values.eigenverbrauch.toFixed(3)}`,
    );

    return values;
  } catch (error) {
    logError(`Fehler bei der Berechnung der aktuellen Werte: ${error}`);
    return null;
  }
};

/**
 * ZWEITE REGEL:
 * Hält zwei Summen-Items (kWh) aktuell:
 *  - PV_Sum_BatteryCharge: gesamte in die Batterie geladene Energie
 *  - PV_Sum_BatteryDischarge: gesamte aus der Batterie entladene Energie
 *
 * Trigger: jede Änderung von PV_Battery_Power (W, < 0 = Laden, > 0 = Entladen)
 *
 * Integration: bei jedem Change Δt zum vorherigen Change, Energie = |P_prev| * Δt
 * Der letzte Messpunkt wird bewusst im privaten Cache gehalten, damit kein
 * überlappendes oder fehlerhaftes Intervall aus der Persistence erneut integriert wird.
 */
const batterySumRule = rules.JSRule({
  id: "PVBatterySumUpdate",
  name: "PV-Batterie Summenaktualisierung",
  description:
    "Aktualisiert PV_Sum_BatteryCharge und PV_Sum_BatteryDischarge aus PV_Battery_Power",
  triggers: [
    triggers.SystemStartlevelTrigger(100),
    triggers.ItemStateChangeTrigger(SOURCE_ITEMS.batteryPower),
  ],
  execute: (event, ctx) => {
    const ruleId = ctx && ctx.ruleUID ? ctx.ruleUID : "PVBatterySumUpdate";
    const logger = LoggerFactory.getLogger("LH30." + ruleId);
    RULE_LOGGER = logger;
    try {
      const batteryItem = items.getItem(SOURCE_ITEMS.batteryPower);
      const chargeItem = items.getItem(SOURCE_ITEMS.batteryChargeSum);
      const dischargeItem = items.getItem(SOURCE_ITEMS.batteryDischargeSum);

      const cacheKey = "pvBatteryAccumulatorState";
      const nowMs = JavaZDT.now(ZoneId.of(TIMEZONE)).toInstant().toEpochMilli();

      const parseNumber = (stateObj) => {
        if (
          !stateObj ||
          stateObj.toString() === "NULL" ||
          stateObj.toString() === "UNDEF"
        ) {
          return null;
        }

        const parsed = parseFloat(stateObj.toString().split(" ")[0]);
        return isNaN(parsed) ? null : parsed;
      };

      const parseEnergyKWhStrict = (itemName, stateObj) => {
        if (!stateObj) {
          logWarn(
            `${itemName} has no state object. Battery accumulator update skipped.`,
          );
          return null;
        }

        const stateText = stateObj.toString().trim();

        if (!stateText || stateText === "NULL" || stateText === "UNDEF") {
          logWarn(
            `${itemName} has invalid state '${stateText || "empty"}'. Battery accumulator update skipped.`,
          );
          return null;
        }

        const match = stateText.match(
          /^([-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)\s*([a-zA-Z]+)?$/,
        );

        if (!match) {
          logWarn(
            `${itemName} has unparsable state '${stateText}'. Battery accumulator update skipped.`,
          );
          return null;
        }

        const value = Number(match[1]);
        const unit = match[2] || "kWh";

        if (!Number.isFinite(value)) {
          logWarn(
            `${itemName} has non-finite state '${stateText}'. Battery accumulator update skipped.`,
          );
          return null;
        }

        if (unit === "Wh") {
          return value / 1000.0;
        }

        if (unit === "kWh") {
          return value;
        }

        if (unit === "MWh") {
          return value * 1000.0;
        }

        logWarn(
          `${itemName} has unsupported energy unit '${unit}' in state '${stateText}'. Battery accumulator update skipped.`,
        );
        return null;
      };

      const getLastPersistedEnergyKWh = (itemName) => {
        try {
          const item = items.getItem(itemName);
          if (!item) {
            logWarn(
              `Could not retrieve persisted fallback for ${itemName}: item not found.`,
            );
            return null;
          }

          let historicItem = null;

          const persistenceApi = item.persistence;
          if (!persistenceApi) {
            logDebug(
              `Item persistence API unavailable for ${itemName}.`,
            );
          } else {
            try {
              historicItem = persistenceApi.previousState(
                true,
                "influxdb",
              );
            } catch (e) {
              logDebug(
                `item.persistence.previousState(true) not available for ${itemName}: ${e}`,
              );
            }

            if (!historicItem) {
              try {
                historicItem = persistenceApi.previousState(
                  false,
                  "influxdb",
                );
              } catch (e) {
                logDebug(
                  `item.persistence.previousState(false) not available for ${itemName}: ${e}`,
                );
              }
            }

            if (!historicItem) {
              try {
                historicItem = persistenceApi.persistedState(
                  JavaZDT.now(ZoneId.of(TIMEZONE)),
                  "influxdb",
                );
              } catch (e) {
                logDebug(
                  `item.persistence.persistedState(now) lookup failed for ${itemName}: ${e}`,
                );
              }
            }
          }

          if (!historicItem) {
            logDebug(`No persisted fallback state found for ${itemName}`);
            return null;
          }

          const persistedState =
            historicItem.state !== undefined
              ? historicItem.state
              : historicItem.getState?.();
          const parsed = parseEnergyKWhStrict(itemName, persistedState);

          if (parsed !== null) {
            logInfo(
              `Using restored persisted state for ${itemName}: ${persistedState.toString()}`,
            );
            if (
              !item.state ||
              item.state.toString() === "NULL" ||
              item.state.toString() === "UNDEF"
            ) {
              item.postUpdate(`${parsed.toFixed(4)} kWh`);
            }
            return parsed;
          }

          logWarn(
            `Persisted fallback state for ${itemName} is invalid: ${persistedState.toString()}`,
          );
        } catch (e) {
          logError(`Error loading persisted fallback for ${itemName}: ${e}`);
        }
        return null;
      };

      const getEnergyKWhOrPersistedFallback = (itemName, stateObj) => {
        const parsed = parseEnergyKWhStrict(itemName, stateObj);
        if (parsed !== null) {
          return parsed;
        }
        logWarn(
          `${itemName} state invalid, attempting to restore last persisted value from history.`,
        );
        return getLastPersistedEnergyKWh(itemName);
      };

      const currentPower = parseNumber(batteryItem.state);
      if (currentPower === null) {
        logDebug(
          `battery power item ${SOURCE_ITEMS.batteryPower} has no usable numeric state`,
        );
        return;
      }

      const previous = cache.private.get(cacheKey);

      if (
        !previous ||
        (event && event.triggerType === "SystemStartlevelTrigger")
      ) {
        cache.private.put(cacheKey, {
          timestampMs: nowMs,
          powerW: currentPower,
        });
        logDebug(
          `Initialized battery accumulator cache: power=${currentPower.toFixed(1)}W`,
        );
        return;
      }

      const dtHours = (nowMs - previous.timestampMs) / 3600000.0;
      if (dtHours <= 0) {
        cache.private.put(cacheKey, {
          timestampMs: nowMs,
          powerW: currentPower,
        });
        return;
      }

      const previousPower = previous.powerW;
      const energyKWh = (Math.abs(previousPower) * dtHours) / 1000.0;

      const chargeSumParsed = getEnergyKWhOrPersistedFallback(
        SOURCE_ITEMS.batteryChargeSum,
        chargeItem.state,
      );
      const dischargeSumParsed = getEnergyKWhOrPersistedFallback(
        SOURCE_ITEMS.batteryDischargeSum,
        dischargeItem.state,
      );

      if (chargeSumParsed === null || dischargeSumParsed === null) {
        logWarn(
          `Battery accumulator skipped because ${SOURCE_ITEMS.batteryChargeSum} or ` +
            `${SOURCE_ITEMS.batteryDischargeSum} is invalid or not restored yet. No sum item will be written.`,
        );

        cache.private.put(cacheKey, {
          timestampMs: nowMs,
          powerW: currentPower,
        });
        return;
      }

      let chargeSum = chargeSumParsed;
      let dischargeSum = dischargeSumParsed;

      if (energyKWh > 0) {
        if (previousPower < 0) {
          chargeSum += energyKWh;
          chargeItem.postUpdate(`${chargeSum.toFixed(4)} kWh`);
          logDebug(
            `battery charge: dt=${dtHours.toFixed(4)} h, ` +
              `power=${previousPower.toFixed(1)} W, ` +
              `+${energyKWh.toFixed(4)} kWh → sum=${chargeSum.toFixed(4)} kWh`,
          );
        } else if (previousPower > 0) {
          dischargeSum += energyKWh;
          dischargeItem.postUpdate(`${dischargeSum.toFixed(4)} kWh`);
          logDebug(
            `battery discharge: dt=${dtHours.toFixed(4)} h, ` +
              `power=${previousPower.toFixed(1)} W, ` +
              `+${energyKWh.toFixed(4)} kWh → sum=${dischargeSum.toFixed(4)} kWh`,
          );
        } else {
          logDebug(
            `battery power is 0 W for dt=${dtHours.toFixed(4)} h, no energy added`,
          );
        }
      }

      cache.private.put(cacheKey, {
        timestampMs: nowMs,
        powerW: currentPower,
      });
    } catch (error) {
      logError(`Fehler in der Batterie-Summenregel: ${error}`);
    }
  },
});

// Hauptregel für die Berechnungen
const pvRule = rules.JSRule({
  id: "PVPlantSumUpdate",
  name: "PV-Anlage Summenaktualisierung",
  description: "Berechnet PV-Anlage Summenwerte aus Persistence",
  triggers: [
    triggers.SystemStartlevelTrigger(100),
    triggers.GenericCronTrigger("5 */5 * * * ?"), // Alle 5 Minuten
    triggers.GenericCronTrigger("35 1 0 * * ?"), // 1 Minute nach Mitternacht
    triggers.ItemStateChangeTrigger("PV_Zeitraum"),
  ],
  execute: async (event, ctx) => {
    const ruleId = ctx && ctx.ruleUID ? ctx.ruleUID : "PVPlantSumUpdate";
    const logger = LoggerFactory.getLogger("LH30." + ruleId);
    RULE_LOGGER = logger;
    try {
      // Sofort anzeigen, dass gerechnet wird
      items.getItem("PV_Erzeugung").postUpdate("(Berechnung...)");
      items.getItem("PV_Gesamtverbrauch").postUpdate("(Berechnung...)");

      items.getItem("PV_Eigenverbrauch").postUpdate("--");
      items.getItem("PV_Direktverbrauch").postUpdate("--");
      items.getItem("PV_Batterieladung").postUpdate("--");
      items.getItem("PV_Netzeinspeisung").postUpdate("--");
      items.getItem("PV_Eigenversorgung").postUpdate("--");
      items.getItem("PV_EigenversorgungPV").postUpdate("--");
      items.getItem("PV_EigenversorgungBatterie").postUpdate("--");
      items.getItem("PV_Netzbezug").postUpdate("--");

      if (
        event &&
        event.triggerType === "GenericCronTrigger" &&
        event.triggerConfig === "0 1 0 * * ?"
      ) {
        await new Promise((resolve) => setTimeout(resolve, 5000));
      }

      const zeitraum = getConfiguredZeitraum();
      const [startTime, endTime] = getTimeRange(zeitraum);

      logDebug(
        `event=${event ? event.itemName || event.triggerType || "unknown" : "manual"}, ` +
          `zeitraum=${zeitraum}, ` +
          `start=${formatDateTime(startTime.toInstant().toEpochMilli())}, ` +
          `end=${formatDateTime(endTime.toInstant().toEpochMilli())}`,
      );

      const values = await calculateCurrentValues(startTime, endTime);

      if (values) {
        logPvSummary({
          zeitraum,
          startTime,
          endTime,
          values,
        });
        updateDisplayItems(values);
        logInfo(`Werte für ${zeitraum} aktualisiert`);
      }
    } catch (error) {
      logError(`Fehler in der PV-Regel: ${error}`);
    }
  },
});
