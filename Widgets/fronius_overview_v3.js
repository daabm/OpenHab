// https://community.openhab.org/t/fronius-widget-for-generation-and-consumption-including-rules/162609
const { items, rules, time, actions, triggers } = require('openhab');

// Java-Zeitklassen importieren
const JavaZDT = Java.type('java.time.ZonedDateTime');
const ZoneId = Java.type('java.time.ZoneId');
const ChronoField = Java.type('java.time.temporal.ChronoField');
const Instant = Java.type('java.time.Instant');

// Konstanten
const TIMEZONE = 'Europe/Berlin';
const MINIMAL_POWER_THRESHOLD = 0.001; // 1 Wh Schwelle

// gridDraw:            PV_Meter_EnergyConsumed (kWh, kumuliert)
// gridFeed:            PV_Meter_EnergyProduced (kWh, kumuliert)
// load:                PV_Inverter1_Power (W)
// batteryPower:        PV_Battery_Power (W, < 0 = Laden, > 0 = Entladen)
// production:          PV_Production_Power (W)
// NEW:                 PV_Sum_BatteryCharge (kWh, kumuliert)
//                      PV_Sum_BatteryDischarge (kWh, kumuliert)
const SOURCE_ITEMS = {
    gridDraw: 'PV_Meter_EnergyConsumed',
    gridFeed: 'PV_Meter_EnergyProduced',
    load: 'PV_Load_Power',
    batteryPower: 'PV_Battery_Power',
    production: 'PV_Production_Power',
    batteryChargeSum: 'PV_Sum_BatteryCharge',
    batteryDischargeSum: 'PV_Sum_BatteryDischarge'
};

// Logger-Funktionen
const log = (message, level = 'info') => {
    // const advancedLogging = items.getItem('Advanced_Logging').state === 'ON';
    const advancedLogging = false;
    if (advancedLogging || level === 'error') {
        switch (level) {
            case 'debug': console.log('[DEBUG] ' + message); break;
            case 'warn': console.warn(message); break;
            case 'error': console.error(message); break;
            default: console.log(message);
        }
    }
};

const formatDateTime = timestamp => {
    const date = new Date(timestamp);
    return date.toLocaleString('de-DE', {
        timeZone: TIMEZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
};

// Zeit-Funktionen
const getTimeRange = (zeitraum) => {
    const zoneId = ZoneId.of(TIMEZONE);
    const now = JavaZDT.now(zoneId);
    const dataCollectionStart = JavaZDT.of(2026, 1, 1, 0, 0, 0, 0, zoneId);

    const startOfDay = (date) => date.withHour(0).withMinute(0).withSecond(0).withNano(0);
    const endOfDay = (date) => date.withHour(23).withMinute(59).withSecond(59).withNano(999999999);

    let startTime, endTime;

    switch (zeitraum.toLowerCase()) {
        case 'heute':
            startTime = startOfDay(now);
            endTime = now;
            break;
        case 'gestern':
            startTime = startOfDay(now.minusDays(1));
            endTime = endOfDay(now.minusDays(1));
            break;
        case 'woche':
            // AKTUELLE Woche (Montag–heute), nicht letzte 7 Tage
            startTime = startOfDay(
                now.with(ChronoField.DAY_OF_WEEK, 1) // Montag dieser Woche
            );
            endTime = now;
            break;
        case 'monat':
            startTime = startOfDay(now.withDayOfMonth(1));
            endTime = now;
            break;
        case 'jahr':
            startTime = startOfDay(now.withDayOfYear(1));
            endTime = now;
            break;
        default:
            startTime = startOfDay(now);
            endTime = now;
    }

    log(`Ursprünglicher Zeitraum für ${zeitraum}:`, 'debug');
    log(`  Start: ${formatDateTime(startTime.toInstant().toEpochMilli())}`, 'debug');
    log(`  Ende:  ${formatDateTime(endTime.toInstant().toEpochMilli())}`, 'debug');

    if (startTime.isBefore(dataCollectionStart)) {
        startTime = dataCollectionStart;
        log(`Startzeit auf Datenerfassungsbeginn angepasst: ${formatDateTime(startTime.toInstant().toEpochMilli())}`, 'debug');
    }

    log(`Finaler Zeitraum ${zeitraum}:`, 'debug');
    log(`  Start: ${formatDateTime(startTime.toInstant().toEpochMilli())}`, 'debug');
    log(`  Ende:  ${formatDateTime(endTime.toInstant().toEpochMilli())}`, 'debug');

    return [startTime, endTime];
};

// Hilfsfunktion zum Aktualisieren der Anzeige-Items
const updateDisplayItems = (values) => {
    try {
        log('Starte Update der Display-Items mit folgenden Werten:', 'debug');
        Object.entries(values).forEach(([key, value]) => {
            log(`${key}: ${value}`, 'debug');
        });

        const formatValue = value => value < 1
            ? `${(value * 1000).toFixed(1)} Wh`
            : `${value.toFixed(2)} kWh`;

        items.getItem('PV_Erzeugung').postUpdate(formatValue(values.erzeugung));
        items.getItem('PV_Eigenverbrauch').postUpdate(formatValue(values.eigenverbrauch));
        items.getItem('PV_Direktverbrauch').postUpdate(formatValue(values.direktverbrauch));
        items.getItem('PV_Batterieladung').postUpdate(formatValue(values.batterieladung));
        items.getItem('PV_Netzeinspeisung').postUpdate(formatValue(values.netzeinspeisung));
        items.getItem('PV_Gesamtverbrauch').postUpdate(formatValue(values.gesamtverbrauch));
        items.getItem('PV_Eigenversorgung').postUpdate(formatValue(values.eigenversorgung));
        items.getItem('PV_EigenversorgungPV').postUpdate(formatValue(values.eigenversorgungPV));
        items.getItem('PV_EigenversorgungBatterie').postUpdate(formatValue(values.eigenversorgungBatterie));
        items.getItem('PV_Netzbezug').postUpdate(formatValue(values.netzbezug));

        const prozentWerte = {
            eigenverbrauchProzent: 0,
            direktverbrauchProzent: 0,
            batterieladungProzent: 0,
            netzeinspeisungProzent: 0,
            eigenversorgungProzent: 0,
            eigenversorgungPVProzent: 0,
            eigenversorgungBatterieProzent: 0,
            netzbezugProzent: 0
        };

        if (values.erzeugung > MINIMAL_POWER_THRESHOLD) {
            prozentWerte.eigenverbrauchProzent = Math.round((values.eigenverbrauch / values.erzeugung) * 100);
            prozentWerte.direktverbrauchProzent = Math.round((values.direktverbrauch / values.erzeugung) * 100);
            prozentWerte.batterieladungProzent = Math.round((values.batterieladung / values.erzeugung) * 100);
            prozentWerte.netzeinspeisungProzent = Math.round((values.netzeinspeisung / values.erzeugung) * 100);
        }

        if (values.gesamtverbrauch > MINIMAL_POWER_THRESHOLD) {
            const effectiveNetzbezug = values.netzbezug < MINIMAL_POWER_THRESHOLD ? 0 : values.netzbezug;
            const effectiveBatterie = values.eigenversorgungBatterie;
            const effectivePV = values.eigenversorgungPV;

            let batterieProzent = Math.round((effectiveBatterie / values.gesamtverbrauch) * 100);
            let pvProzent = Math.round((effectivePV / values.gesamtverbrauch) * 100);
            let netzbezugProzent = Math.round((effectiveNetzbezug / values.gesamtverbrauch) * 100);

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

        items.getItem('PV_EigenverbrauchProzent').postUpdate(prozentWerte.eigenverbrauchProzent);
        items.getItem('PV_DirektverbrauchProzent').postUpdate(prozentWerte.direktverbrauchProzent);
        items.getItem('PV_BatterieladungProzent').postUpdate(prozentWerte.batterieladungProzent);
        items.getItem('PV_NetzeinspeisungProzent').postUpdate(prozentWerte.netzeinspeisungProzent);
        items.getItem('PV_EigenversorgungProzent').postUpdate(prozentWerte.eigenversorgungProzent);
        items.getItem('PV_EigenversorgungPVProzent').postUpdate(prozentWerte.eigenversorgungPVProzent);
        items.getItem('PV_EigenversorgungBatterieProzent').postUpdate(prozentWerte.eigenversorgungBatterieProzent);
        items.getItem('PV_NetzbezugProzent').postUpdate(prozentWerte.netzbezugProzent);

        log('Display-Items erfolgreich aktualisiert', 'debug');
    } catch (error) {
        log(`Fehler beim Update der Display-Items: ${error}`, 'error');
    }
};

// kWh-Berechnung aus Summen-Items (kWh, kumuliert)
const getEnergyDeltaFromSumItem = (itemName, startTime, endTime) => {
    try {
        const it = items.getItem(itemName);
        const delta = actions.PersistenceExtensions.deltaBetween(it, startTime, endTime, 'influxdb');
        if (!delta || !delta.doubleValue) {
            log(`deltaBetween liefert keinen Wert für ${itemName}`, 'debug');
            return 0;
        }
        const val = delta.doubleValue(); // bereits kWh
        log(`${itemName} deltaBetween: ${val.toFixed(3)} kWh`, 'debug');
        return val < 0 ? 0 : val;
    } catch (e) {
        log(`Fehler bei deltaBetween(${itemName}): ${e}`, 'error');
        return 0;
    }
};

// kWh-Berechnung aus Leistungs-Items in W → Wh → kWh (über Durchschnitt * Dauer)
const getEnergyFromPowerAverageW = (itemName, startTime, endTime) => {
    try {
        const it = items.getItem(itemName);
        const avg = actions.PersistenceExtensions.averageBetween(it, startTime, endTime, 'influxdb');
        if (!avg || !avg.doubleValue) {
            log(`averageBetween liefert keinen Wert für ${itemName}`, 'debug');
            return 0;
        }
        const avgW = avg.doubleValue(); // W
        const durationHours =
            (endTime.toInstant().toEpochMilli() - startTime.toInstant().toEpochMilli()) / 3600000.0;
        const energyWh = avgW * durationHours;
        const energyKWh = energyWh / 1000.0;
        log(`${itemName} averageBetween: ${avgW.toFixed(1)} W über ${durationHours.toFixed(3)} h = ${energyKWh.toFixed(3)} kWh`, 'debug');
        return energyKWh < 0 ? 0 : energyKWh;
    } catch (e) {
        log(`Fehler bei averageBetween(${itemName}): ${e}`, 'error');
        return 0;
    }
};

// kWh-Berechnung für Batterie-Ladung/-Entladung aus den neuen Summen-Items
const getBatteryChargeEnergyFromSum = (startTime, endTime) =>
    getEnergyDeltaFromSumItem(SOURCE_ITEMS.batteryChargeSum, startTime, endTime);

const getBatteryDischargeEnergyFromSum = (startTime, endTime) =>
    getEnergyDeltaFromSumItem(SOURCE_ITEMS.batteryDischargeSum, startTime, endTime);

const calculateCurrentValues = async (startTime, endTime) => {
    try {
        log(`Berechne Werte von ${startTime} bis ${endTime}`, 'debug');

        const gridDraw = getEnergyDeltaFromSumItem(SOURCE_ITEMS.gridDraw, startTime, endTime);
        const gridFeed = getEnergyDeltaFromSumItem(SOURCE_ITEMS.gridFeed, startTime, endTime);

        const production = getEnergyFromPowerAverageW(SOURCE_ITEMS.production, startTime, endTime);
        const load = getEnergyFromPowerAverageW(SOURCE_ITEMS.load, startTime, endTime);

        // NEU: Batterie-Energie aus den Summen-Items (kWh, kumuliert)
        const batteryCharge = getBatteryChargeEnergyFromSum(startTime, endTime);
        const batteryDischarge = getBatteryDischargeEnergyFromSum(startTime, endTime);

        log('Rohe Energiemengen:', 'debug');
        log(`Netzbezug (gridDraw): ${gridDraw.toFixed(3)} kWh`, 'debug');
        log(`Netzeinspeisung (gridFeed): ${gridFeed.toFixed(3)} kWh`, 'debug');
        log(`Produktion (production): ${production.toFixed(3)} kWh`, 'debug');
        log(`Last (load): ${load.toFixed(3)} kWh`, 'debug');
        log(`Batterieladung (batteryCharge): ${batteryCharge.toFixed(3)} kWh`, 'debug');
        log(`Batterieentladung (batteryDischarge): ${batteryDischarge.toFixed(3)} kWh`, 'debug');

        const values = {
            netzbezug: gridDraw,
            netzeinspeisung: gridFeed,
            batterieladung: batteryCharge,
            eigenversorgungBatterie: batteryDischarge,
            direktverbrauch: Math.max(0, production - gridFeed - batteryCharge),
            verbrauch: load
        };

        log('Berechnete Basiswerte:', 'debug');
        Object.entries(values).forEach(([key, value]) =>
            log(`${key}: ${value.toFixed(3)} kWh`, 'debug')
        );

        values.eigenverbrauch = values.direktverbrauch + values.batterieladung;
        values.erzeugung = values.netzeinspeisung + values.eigenverbrauch;
        values.eigenversorgungPV = values.direktverbrauch;
        values.eigenversorgung = values.eigenversorgungPV + values.eigenversorgungBatterie;
        values.gesamtverbrauch = values.eigenversorgung + values.netzbezug;

        log('Abgeleitete Werte:', 'debug');
        Object.entries(values).forEach(([key, value]) =>
            log(`${key}: ${value.toFixed(3)} kWh`, 'debug')
        );

        log(
            `Gesamtverbrauch = Eigenversorgung + Netzbezug: ` +
            `${values.gesamtverbrauch.toFixed(3)} = ` +
            `${values.eigenversorgung.toFixed(3)} + ${values.netzbezug.toFixed(3)}`,
            'debug'
        );
        log(
            `Erzeugung = Netzeinspeisung + Eigenverbrauch: ` +
            `${values.erzeugung.toFixed(3)} = ` +
            `${values.netzeinspeisung.toFixed(3)} + ${values.eigenverbrauch.toFixed(3)}`,
            'debug'
        );

        return values;
    } catch (error) {
        log(`Fehler bei der Berechnung der aktuellen Werte: ${error}`, 'error');
        return null;
    }
};

// Hauptregel für die Berechnungen
const pvRule = rules.JSRule({
    name: "PV-Anlage Berechnung",
    description: "Berechnet PV-Anlagen Werte aus Persistence (mit Batteriesummen)",
    triggers: [
        triggers.SystemStartlevelTrigger(100),
        triggers.GenericCronTrigger("0 */5 * * * ?"),    // Alle 5 Minuten
        triggers.GenericCronTrigger("0 1 0 * * ?"),      // 1 Minute nach Mitternacht
        triggers.ItemStateChangeTrigger("PV_Zeitraum")
    ],
    execute: async (event) => {
        try {
            // Sofort anzeigen, dass gerechnet wird
            items.getItem('PV_Erzeugung').postUpdate('(Berechnung...)');
            items.getItem('PV_Gesamtverbrauch').postUpdate('(Berechnung...)');

            items.getItem('PV_Eigenverbrauch').postUpdate('--');
            items.getItem('PV_Direktverbrauch').postUpdate('--');
            items.getItem('PV_Batterieladung').postUpdate('--');
            items.getItem('PV_Netzeinspeisung').postUpdate('--');
            items.getItem('PV_Eigenversorgung').postUpdate('--');
            items.getItem('PV_EigenversorgungPV').postUpdate('--');
            items.getItem('PV_EigenversorgungBatterie').postUpdate('--');
            items.getItem('PV_Netzbezug').postUpdate('--');

            if (event.triggerType === 'GenericCronTrigger' &&
                event.triggerConfig === "0 1 0 * * ?") {
                await new Promise(resolve => setTimeout(resolve, 5000));
            }

            const zeitraum = items.getItem('PV_Zeitraum').state;
            const [startTime, endTime] = getTimeRange(zeitraum);
            const values = await calculateCurrentValues(startTime, endTime);

            if (values) {
                updateDisplayItems(values);
                log(`Werte für ${zeitraum} aktualisiert`);
            }
        } catch (error) {
            log(`Fehler in der PV-Regel: ${error}`, 'error');
        }
    }
});

/**
 * ZWEITE REGEL:
 * Hält zwei Summen-Items (kWh) aktuell:
 *  - PV_Sum_BatteryCharge: gesamte in die Batterie geladene Energie
 *  - PV_Sum_BatteryDischarge: gesamte aus der Batterie entladene Energie
 *
 * Trigger: jede Änderung von PV_Battery_Power (W, < 0 = Laden, > 0 = Entladen)
 *
 * Integration: bei jedem Change Δt zum vorherigen Change, Energie = |P_avg| * Δt
 */
const batterySumRule = rules.JSRule({
    name: "PV-Batterie Summenaktualisierung",
    description: "Aktualisiert PV_Sum_BatteryCharge und PV_Sum_BatteryDischarge aus PV_Battery_Power",
    triggers: [
        triggers.ItemStateChangeTrigger(SOURCE_ITEMS.batteryPower)
    ],
    execute: (event) => {
        try {
            const batteryItem = items.getItem(SOURCE_ITEMS.batteryPower);
            const chargeItem = items.getItem(SOURCE_ITEMS.batteryChargeSum);
            const dischargeItem = items.getItem(SOURCE_ITEMS.batteryDischargeSum);

            const persistence = actions.PersistenceExtensions;

            // aktuelle Zeit und letzte State-Änderung
            const nowZdt = JavaZDT.now(ZoneId.of(TIMEZONE));
            const prevHist = persistence.previousState(batteryItem, true, 'influxdb');

            if (!prevHist || !prevHist.state || !prevHist.state.doubleValue) {
                log('Keine vorherigen Batteriewerte verfügbar – Summenstart.', 'debug');
                return;
            }

            const prevTime = prevHist.getTimestamp();
            const prevMs = prevTime.toInstant().toEpochMilli();
            const nowMs = nowZdt.toInstant().toEpochMilli();
            const dtHours = (nowMs - prevMs) / 3600000.0;
            if (dtHours <= 0) return;

            // Leistung: wir nehmen den alten Wert als konstant im Intervall
            const prevP = prevHist.state.doubleValue(); // W

            // Energie in diesem Intervall (kWh)
            const energyKWh = Math.abs(prevP) * dtHours / 1000.0;
            if (energyKWh <= 0) return;

            // Aktuelle Summen aus Items lesen
            const parseSum = (stateObj) => {
                if (!stateObj || stateObj.toString() === 'NULL' || stateObj.toString() === 'UNDEF') {
                    return 0;
                }
                return parseFloat(stateObj.toString().split(' ')[0]) || 0;
            };

            let chargeSum = parseSum(chargeItem.state);
            let dischargeSum = parseSum(dischargeItem.state);

            if (prevP < 0) {
                // Laden
                chargeSum += energyKWh;
                chargeItem.postUpdate(chargeSum.toFixed(4));
                log(`Batterie-Laden: +${energyKWh.toFixed(4)} kWh → Summe=${chargeSum.toFixed(4)} kWh`, 'debug');
            } else if (prevP > 0) {
                // Entladen
                dischargeSum += energyKWh;
                dischargeItem.postUpdate(dischargeSum.toFixed(4));
                log(`Batterie-Entladen: +${energyKWh.toFixed(4)} kWh → Summe=${dischargeSum.toFixed(4)} kWh`, 'debug');
            }
        } catch (error) {
            log(`Fehler in der Batterie-Summenregel: ${error}`, 'error');
        }
    }
});
