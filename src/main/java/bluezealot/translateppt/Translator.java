package bluezealot.translateppt;

import java.io.IOException;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;

/**
 * Call openai api to translate
 * From given language to target language
 */
public class Translator {
    private final String OPENAI_API_KEY = "";
    private final String OPENAI_URL = "https://api.openai.com/v1/chat/completions";
    private final OkHttpClient client = new OkHttpClient();
    private final Gson gson = new Gson();

    private static final int MAX_RETRIES = 3;
    private static final int INITIAL_DELAY_MS = 1000;

    private String makeApiCall(Request request) throws IOException {
        int delay = INITIAL_DELAY_MS;
        for (int attempt = 1; attempt <= MAX_RETRIES; attempt++) {
            try {
                try (Response response = client.newCall(request).execute()) {
                    if (!response.isSuccessful()) {
                        throw new IOException("Unexpected response: " + response);
                    }
                    return response.body().string();
                }
            } catch (Exception e) {
                if (attempt == MAX_RETRIES) {
                    throw e;
                }
                System.out.println("API call failed. Attempt " + attempt + " of " + MAX_RETRIES + ". Retrying in " + delay + "ms...");
                try {
                    Thread.sleep(delay);
                } catch (InterruptedException ie) {
                    Thread.currentThread().interrupt();
                    throw new IOException("Thread interrupted during retry wait", ie);
                }
                delay *= 2; // Double the delay for next attempt
            }
        }
        return null; // Should never reach here
    }

    public String translate(String sourceText, String sourceLang, String targetLang) throws IOException {
        JsonObject message1 = new JsonObject();
        message1.addProperty("role", "system");
        message1.addProperty("content", "You are a translation engine. Translate the user's text from " + sourceLang + " to " + targetLang + ".");

        JsonObject message2 = new JsonObject();
        message2.addProperty("role", "user");
        message2.addProperty("content", sourceText);

        JsonArray messages = new JsonArray();
        messages.add(message1);
        messages.add(message2);

        JsonObject requestBody = new JsonObject();
        requestBody.addProperty("model", "gpt-4.1-mini-2025-04-14");
        requestBody.add("messages", messages);

        Request request = new Request.Builder()
                .url(OPENAI_URL)
                .addHeader("Authorization", "Bearer " + OPENAI_API_KEY)
                .addHeader("Content-Type", "application/json")
                .post(RequestBody.create(gson.toJson(requestBody), MediaType.get("application/json")))
                .build();

        String responseString = makeApiCall(request);
        JsonObject responseJson = JsonParser.parseString(responseString).getAsJsonObject();
        return responseJson
                .getAsJsonArray("choices")
                .get(0).getAsJsonObject()
                .getAsJsonObject("message")
                .get("content").getAsString().trim();
    }
}