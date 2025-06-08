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

        try (Response response = client.newCall(request).execute()) {
            if (!response.isSuccessful()) {
                throw new IOException("Unexpected response: " + response);
            }

            JsonObject responseJson = JsonParser.parseString(response.body().string()).getAsJsonObject();
            return responseJson
                    .getAsJsonArray("choices")
                    .get(0).getAsJsonObject()
                    .getAsJsonObject("message")
                    .get("content").getAsString().trim();
        }
    }
}
