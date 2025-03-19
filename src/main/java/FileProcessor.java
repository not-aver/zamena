import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.sax.BodyContentHandler;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FileProcessor {
    private static final HttpClient client = HttpClient.newHttpClient();
    private static final String SCHEDULE_URL = "https://vgpgk.ru/raspisanie/1-korpus_1-smena_1-semestr_2022.xls?v=2023011141816";
    private static final String REPLACEMENTS_URL = "https://vgpgk.ru/raspisanie/vgpgk-zameny-1-korpus.doc?v=2023011141816";
    private static final Logger logger = LoggerFactory.getLogger(FileProcessor.class);
    private static final Pattern GROUP_PATTERN = Pattern.compile("(И[А-Я]{1,2}-\\d{3}(?:\\s*\\([^)]+\\))?)|(К[А-Я]{1,2}-\\d{3})|(С[А-Я]{1,2}-\\d{3})|(ПС-\\d{3})");
    private static final Pattern REPLACEMENT_PATTERN = Pattern.compile("^\\s*(\\d\\sп\\..*)$");
    private static final Pattern SPECIAL_REPLACEMENT_PATTERN = Pattern.compile("^(УП\\..*)$");
    private static final Pattern SHIFT_PATTERN = Pattern.compile("^\\d\\sкорпус\\s\\d\\sсмена$");
    private static final Pattern DEPARTMENT_PATTERN = Pattern.compile("^Отделение\\s[А-ЯА-я]+");

    public static XSSFWorkbook downloadAndConvertSchedule() throws Exception {
        logger.info("Скачивание файла расписания с URL: {}", SCHEDULE_URL);
        InputStream xlsStream = downloadFile(SCHEDULE_URL);
        HSSFWorkbook oldWorkbook = new HSSFWorkbook(xlsStream);
        XSSFWorkbook newWorkbook = new XSSFWorkbook();
        newWorkbook.createSheet(oldWorkbook.getSheetName(0));
        logger.info("Конвертация .xls в .xlsx");
        logger.info("Конвертация завершена, листов: {}", newWorkbook.getNumberOfSheets());
        return newWorkbook;
    }

    public static String downloadAndExtractReplacementsText() throws Exception {
        logger.info("Скачивание файла замен с URL: {}", REPLACEMENTS_URL);
        InputStream docStream = downloadFile(REPLACEMENTS_URL);
        return extractTextFromDoc(docStream);
    }

    private static InputStream downloadFile(String urlString) throws Exception {
        HttpRequest request = HttpRequest.newBuilder().uri(new URI(urlString)).GET().build();
        HttpResponse<InputStream> response = client.send(request, HttpResponse.BodyHandlers.ofInputStream());
        return response.body();
    }

    private static String extractTextFromDoc(InputStream docStream) throws Exception {
        logger.info("Извлечение текста из .doc");
        BodyContentHandler handler = new BodyContentHandler();
        Metadata metadata = new Metadata();
        new AutoDetectParser().parse(docStream, handler, metadata, new ParseContext());
        String text = handler.toString();
        logger.info("Текст успешно извлечен, первые 200 символов: {}", text.substring(0, Math.min(200, text.length())));
        return text;
    }

    private static String formatReplacement(String replacement) {
        String[] parts = replacement.trim().split("\\s+", 3);
        if (parts.length < 2) return replacement;

        String pair = parts[0] + " " + parts[1]; // "X п."
        String rest = parts.length > 2 ? parts[2] : "";
        String[] restParts = rest.split(",\\s*");

        if (restParts.length >= 2) {
            String subject = restParts[0];
            String teacherRoom = "(" + restParts[1];
            if (restParts.length > 2) teacherRoom += ", " + restParts[2];
            teacherRoom += ")";
            return pair + " " + subject + " " + teacherRoom;
        } else if (restParts.length == 1 && rest.matches(".*\\d{3}")) {
            return pair + " (" + rest + ")";
        } else if (rest.equals("нет")) {
            return pair + " нет";
        }
        return pair + " " + rest; // Для случаев вроде "1 п. Охрана"
    }

    public static Map<String, List<String>> parseReplacements(String text) {
        Map<String, List<String>> replacements = new HashMap<>();
        String[] lines = text.split("\\n");
        List<String> currentGroups = new ArrayList<>();
        List<String> currentReplacements = new ArrayList<>();
        StringBuilder currentReplacement = new StringBuilder();

        for (String line : lines) {
            line = line.trim();
            if (line.isEmpty()) continue;

            Matcher shiftMatcher = SHIFT_PATTERN.matcher(line);
            Matcher deptMatcher = DEPARTMENT_PATTERN.matcher(line);
            Matcher groupMatcher = GROUP_PATTERN.matcher(line);
            Matcher replacementMatcher = REPLACEMENT_PATTERN.matcher(line);
            Matcher specialMatcher = SPECIAL_REPLACEMENT_PATTERN.matcher(line);
            Matcher practiceMatcher = Pattern.compile("^Практика$").matcher(line);

            // Новый блок (смена или отделение)
            if (shiftMatcher.matches() || deptMatcher.matches()) {
                if (!currentGroups.isEmpty() && !currentReplacements.isEmpty()) {
                    distributeReplacements(replacements, currentGroups, currentReplacements);
                }
                currentGroups.clear();
                currentReplacements.clear();
                currentReplacement.setLength(0);
                continue;
            }

            // Практика
            if (practiceMatcher.matches()) {
                if (!currentGroups.isEmpty()) {
                    replacements.computeIfAbsent(currentGroups.get(currentGroups.size() - 1), k -> new ArrayList<>()).add("Практика");
                }
                continue;
            }

            // Группы
            List<String> groupsInLine = new ArrayList<>();
            while (groupMatcher.find()) {
                groupsInLine.add(groupMatcher.group());
            }
            if (!groupsInLine.isEmpty()) {
                if (!currentGroups.isEmpty() && !currentReplacements.isEmpty()) {
                    distributeReplacements(replacements, currentGroups, currentReplacements);
                }
                currentGroups.clear();
                currentReplacements.clear();
                currentReplacement.setLength(0);
                currentGroups.addAll(groupsInLine);
                for (String group : groupsInLine) {
                    replacements.putIfAbsent(group, new ArrayList<>());
                }
                continue;
            }

            // Специальная замена (УП.*)
            if (specialMatcher.matches()) {
                if (currentReplacement.length() > 0) {
                    currentReplacements.add(formatReplacement(currentReplacement.toString()));
                }
                currentReplacement.setLength(0);
                currentReplacements.add(specialMatcher.group(1));
                continue;
            }

            // Обычная замена
            if (replacementMatcher.matches()) {
                if (currentReplacement.length() > 0) {
                    currentReplacements.add(formatReplacement(currentReplacement.toString()));
                }
                currentReplacement.setLength(0);
                currentReplacement.append(replacementMatcher.group(1));
            } else if (currentReplacement.length() > 0) {
                currentReplacement.append(" ").append(line);
            }
        }

        // Добавляем последнюю замену
        if (currentReplacement.length() > 0) {
            currentReplacements.add(formatReplacement(currentReplacement.toString()));
        }

        // Распределяем оставшиеся замены
        if (!currentGroups.isEmpty() && !currentReplacements.isEmpty()) {
            distributeReplacements(replacements, currentGroups, currentReplacements);
        }

        logger.info("Парсинг завершен, найдено групп: {}", replacements.size());
        return replacements;
    }

    private static void distributeReplacements(Map<String, List<String>> replacements, List<String> groups, List<String> replacementsList) {
        for (int i = 0; i < Math.min(groups.size(), replacementsList.size()); i++) {
            replacements.get(groups.get(i)).add(replacementsList.get(i));
        }
    }

    public static void main(String[] args) {
        try {
            logger.info("Проверка SLF4J: {}", LoggerFactory.getLogger(FileProcessor.class).getClass().getName());
            Class.forName("org.apache.poi.hwpf.OldWordFileFormatException");
            logger.info("poi-scratchpad доступен");

            XSSFWorkbook schedule = downloadAndConvertSchedule();
            String replacementsText = downloadAndExtractReplacementsText();
            Map<String, List<String>> replacements = parseReplacements(replacementsText);

            System.out.println("Расписание: " + schedule.getNumberOfSheets() + " листов");
            System.out.println("Замены по группам:");
            replacements.forEach((group, replList) -> {
                System.out.println("Группа: " + group);
                replList.forEach(repl -> System.out.println("- " + repl));
            });
        } catch (Exception e) {
            logger.error("Ошибка при выполнении: ", e);
            e.printStackTrace();
        }
    }
}