# AGENTS.md

Single-module Maven project: `com.gsdd:xls-util:3.0.1` (jar). Builds with the JDK on `PATH` (bytecode targets Java 25); no Maven wrapper. Inherits base plugins/deps from the external parent `com.gsdd:gsdd-parent:1.0.7`.

## Commands

- `mvn test` — run unit tests (surefire, `*Test.java`)
- `mvn verify` — unit + integration tests (failsafe, `It*.java`) + jacoco
- `-Dskip.unit.tests=true` / `-Dskip.integration.tests=true` to skip each suite
- `mvn package` — needed before running Sonar (populates `target/dependency/` via `copy-dependencies`)
- `mvn spotless:apply` — manual format; checkstyle is separate (`mvn checkstyle:check`)

## Gotchas

- **Spotless auto-applies on `validate`** (every build): google-java-format 1.21.0, `ratchetFrom origin/main` so only files changed on the branch are formatted. It applies, it never fails the build. Keep code in google-java-format style; matches `*.properties` and `*.xml` too.
- **OWASP dependency-check never fails the build**: `failBuildOnCVSS=11`. Reports vulnerabilities but passes.
- `.mvn/jvm.config` carries `--add-exports` flags required for Lombok annotation processing on JDK 9+; do not remove them.
- `**/constants/**` is excluded from jacoco coverage and Sonar.
- Sonar requires a local server (`sonar.host.url=http://localhost:9000`) and `SONAR_LOGIN_TOKEN` env var; `sonar-project.properties` references `target/dependency/lombok-*.jar` so run `mvn package` first.

## Code style

- Utility classes via Lombok `@UtilityClass` (final class, private ctor); logging via `@Slf4j`. `com.gsdd.constants.XlsConstants` holds string constants (`.xls`, `.xlsx`).
- Apache POI 5.5.1: `.xlsx` → `XSSFWorkbook`, `.xls` → `HSSFWorkbook`, selected by file-extension suffix.
- Tests: JUnit 5 + Mockito; workbook fixtures live in `src/test/resources` (`test.xls`, `test.xlsx`).
