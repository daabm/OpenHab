You now have **two separate issues**:

1. **Real rule UID collision**  
   Your down-cycle script is trying to register a rule with the up-cycle UID:

   ```text
   a rule with same UID [rs_up_timer_patterns_configurable_up_cycle] already exists
   ```

2. **Actual logging duplication**  
   Since openHAB framework loggers like `ScriptEngineManagerImpl` and `AbstractScriptFileWatcher` are duplicated too, this is not only your JS rule executing twice.

---

## 1. Fix the rule UID collision first

Your `rs_down_timer_patterns.js` likely contains a copied `id` from the up-cycle rule.

Use a different stable UID for the down rule.

You can use the Apply Button on the code block, or switch to Agent Mode to make the update automatically.

```javascript Config/automation/js/rs_down_timer_patterns.js
// ... existing code ...

rules.JSRule({
  id: "rs_down_timer_patterns_configurable_down_cycle",
  name: "RS close gradually - configurable down cycle",
  description:
    "Starts configured floor-level rollershutter closing cycles using per-floor down-cycle configuration items.",
  triggers: [
    // ... existing triggers ...
  ],
  execute: (event) => {
    // ... existing execution code ...
  },
});

// ... rest of code ...
```

Your up rule should keep its own unique UID:

```javascript Config/automation/js/rs_up_timer_patterns.js
// ... existing code ...

rules.JSRule({
  id: "rs_up_timer_patterns_configurable_up_cycle",
  name: "RS open gradually - configurable up cycle",
  description:
    "Starts configured floor-level rollershutter opening cycles using per-floor up-cycle configuration items.",
  triggers: [triggers.GenericCronTrigger(UP_CYCLE_CHECK_CRON)],
  execute: () => {
    startConfiguredUpPatternsForCurrentMinute("minute-check");
  },
});

// ... rest of code ...
```

That fixes the explicit rule-registration error, but not necessarily the duplicated log lines.

---

## 2. Determine whether two openHAB processes write to the same log

This is now the most important check.

Run this on the host:

```bash diagnostics/check-openhab-log-writers.sh
ps -ef | grep -E '[o]penhab|[k]araf|[j]ava'

echo
echo "Processes writing openhab.log:"
lsof /var/log/openhab/openhab.log 2>/dev/null || lsof /opt/openhab/userdata/logs/openhab.log 2>/dev/null
```

You should see **one** Java/openHAB process writing `openhab.log`.

If you see two Java processes with the log file open, that explains everything.

Common causes:

```text
systemd openHAB service + manually started openHAB
systemd openHAB service + Docker container
old openHAB instance still running after migration
two containers mounting the same userdata/log directory
```

Check especially:

```bash diagnostics/check-openhab-services.sh
systemctl status openhab
docker ps
```

If you run openHAB in Docker, make sure the host systemd service is not also running.

---

## 3. Test if Log4j itself duplicates messages

From the Karaf console, run:

```text console
log:log INFO duplicate-test-karaf-12345
```

Then check `openhab.log`.

### If the test line appears once

Then Log4j routing is probably okay, and the duplicates are likely caused by:

```text
two openHAB processes
duplicate script loading
duplicate runtime registration
```

### If the test line appears twice

Then the duplication is inside the logging pipeline.

Given your current `log4j2.xml`, if only one openHAB process is running, this line **should appear once**.

---

## 4. Your current `log4j2.xml` is mostly not the obvious cause

These lines are important:

```xml Userdata/etc/log4j2.xml
<Logger additivity="false" level="INFO" name="org.openhab">
    <AppenderRef ref="LOGFILE"/>
</Logger>

<Logger additivity="false" level="INFO" name="org.openhab.automation.jsscripting">
    <AppenderRef ref="LOGFILE"/>
</Logger>

<Logger additivity="false" level="INFO" name="org.openhab.core.automation">
    <AppenderRef ref="LOGFILE"/>
</Logger>
```

Because `additivity="false"` is set, log events below these namespaces should not also propagate to `Root`.

So this duplicated framework line:

```text
[ipt.internal.ScriptEngineManagerImpl] - Error during evaluation of script ...
[ipt.internal.ScriptEngineManagerImpl] - Error during evaluation of script ...
```

should not be caused by simple logger inheritance.

That makes these more likely:

```text
two running openHAB instances
same log4j2.xml loaded twice somehow
two appenders writing to same target through runtime config
stale Pax Logging/runtime state requiring a full restart
```

---

## 5. Do a full openHAB restart, not only config reload

After fixing the duplicate rule UID, do a full restart:

```bash diagnostics/restart-openhab.sh
sudo systemctl restart openhab
```

Or, for Docker:

```bash diagnostics/restart-openhab-docker.sh
docker restart <openhab-container-name>
```

Then immediately check:

```bash diagnostics/check-openhab-log-writers.sh
lsof /var/log/openhab/openhab.log 2>/dev/null || lsof /opt/openhab/userdata/logs/openhab.log 2>/dev/null
```

Again, there should be only one writer.

---

## 6. Check whether both `/etc/openhab` and another config path are watched

Your log shows:

```text
/etc/openhab/automation/js/rs_down_timer_patterns.js
```

That is normal for Linux package installations.

But check whether the same file also exists in another mounted/config path, especially if Docker or symlinks are involved:

```bash diagnostics/find-duplicate-scripts.sh
find /etc/openhab /opt/openhab -path '*automation/js*' -type f 2>/dev/null | sort
```

Look for duplicate copies like:

```text
/etc/openhab/automation/js/rs_up_timer_patterns.js
/opt/openhab/conf/automation/js/rs_up_timer_patterns.js
```

or symlink loops.

---

## 7. Check effective logger levels in Karaf

In the Karaf console:

```text console
log:get org.openhab
log:get org.openhab.core.automation
log:get org.openhab.automation.jsscripting
log:get org.openhab.automation.jsscripting.openhab-js
log:get ROOT
```

Expected roughly:

```text
org.openhab                                  INFO
org.openhab.core.automation                 INFO
org.openhab.automation.jsscripting          INFO
ROOT                                        WARN
```

This does not directly show appenders, but it confirms whether Karaf runtime config differs from the file.

---

## 8. Temporarily simplify the logger section

If only one process is running and `log:log INFO duplicate-test-karaf-12345` still appears twice, temporarily simplify the logging configuration.

For diagnosis, remove the extra explicit `org.openhab.automation.jsscripting` and `org.openhab.core.automation` appenders and let them be handled only by `org.openhab`.

You can use the Apply Button on the code block, or switch to Agent Mode to make the update automatically.

```xml Userdata/etc/log4j2.xml
<!-- ... existing code ... -->

<Loggers>
    <Root level="WARN">
        <AppenderRef ref="LOGFILE"/>
    </Root>

    <Logger additivity="false" level="INFO" name="org.openhab">
        <AppenderRef ref="LOGFILE"/>
    </Logger>

    <!-- Keep event log isolated -->
    <Logger additivity="false" level="ERROR" name="openhab.event">
        <AppenderRef ref="EVENT"/>
    </Logger>

    <!-- Keep audit log isolated -->
    <Logger additivity="false" level="INFO" name="org.apache.karaf.jaas.modules.audit">
        <AppenderRef ref="AUDIT"/>
    </Logger>

    <!-- Temporarily comment/remove these during duplicate diagnosis:
    <Logger additivity="false" level="INFO" name="org.openhab.automation.jsscripting">
        <AppenderRef ref="LOGFILE"/>
    </Logger>

    <Logger additivity="false" level="INFO" name="org.openhab.core.automation">
        <AppenderRef ref="LOGFILE"/>
    </Logger>
    -->

    <!-- ... rest of your specific loggers ... -->
</Loggers>

<!-- ... rest of code ... -->
```

Then fully restart openHAB and test again.

But honestly, your current config should not duplicate those lines, so I would first suspect **two writers/processes**.

---

## 9. Quick likely diagnosis

Based on the evidence:

```text
same timestamp
same logger
same message
framework loggers duplicated too
current XML has additivity=false for relevant branches
```

Most likely causes, in order:

1. **Two openHAB instances writing to the same `openhab.log`**
2. **Pax/Log4j runtime state needs a full restart**
3. **Same script watched/loaded twice through duplicate config paths or symlinks**
4. **Actual Log4j appender duplication outside the visible XML**

Fix the down-rule UID first, then run the `lsof` check. That will usually reveal the cause quickly.
