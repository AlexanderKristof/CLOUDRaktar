package hu.raktar.pda;

import android.app.Activity;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.IntentFilter;
import android.content.SharedPreferences;
import android.graphics.Color;
import android.net.Uri;
import android.provider.MediaStore;
import android.os.Build;
import android.os.Bundle;
import android.view.KeyEvent;
import android.view.WindowManager;
import android.webkit.JavascriptInterface;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import androidx.core.content.FileProvider;

import com.google.mlkit.vision.common.InputImage;
import com.google.mlkit.vision.text.TextRecognition;
import com.google.mlkit.vision.text.latin.TextRecognizerOptions;

import org.json.JSONObject;

import java.io.File;
import java.util.ArrayList;

/**
 * Raktar PDA (Zebra TC27): a VPS szerver /pda oldalat jeleniti meg.
 * A vonalkod-olvaso (DataWedge) a beolvasott kodot broadcast intentkent kuldi ide,
 * innen a weboldal window.onScan(kod) fuggvenye kapja meg.
 */
public class MainActivity extends Activity {

    static final String DEFAULT_SERVER = "http://217.144.56.10:5000";
    static final String SCAN_ACTION = "hu.raktar.pda.SCAN";
    static final String DW_PROFILE = "RaktarPDA";

    static final int REQ_OCR_PHOTO = 41;

    private WebView web;
    private SharedPreferences prefs;
    private File ocrFile;

    private final BroadcastReceiver scanReceiver = new BroadcastReceiver() {
        @Override
        public void onReceive(Context context, Intent intent) {
            String data = intent.getStringExtra("com.symbol.datawedge.data_string");
            if (data == null || web == null) return;
            web.evaluateJavascript("window.onScan && window.onScan(" + JSONObject.quote(data.trim()) + ")", null);
        }
    };

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        prefs = getSharedPreferences("raktar", MODE_PRIVATE);
        getWindow().addFlags(WindowManager.LayoutParams.FLAG_KEEP_SCREEN_ON);

        web = new WebView(this);
        web.setBackgroundColor(Color.parseColor("#121212"));
        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setCacheMode(WebSettings.LOAD_NO_CACHE);
        web.addJavascriptInterface(new Bridge(), "RaktarApp");
        // A weboldal alert()/confirm() ablakai is megjelenjenek (alapbol a WebView elnyeli oket)
        web.setWebChromeClient(new WebChromeClient());
        web.setWebViewClient(new WebViewClient() {
            @Override
            public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                if (request.isForMainFrame()) showOfflinePage(String.valueOf(error.getDescription()));
            }
        });
        setContentView(web);

        configureDataWedge();
        loadApp();
    }

    private String server() {
        return prefs.getString("server", DEFAULT_SERVER);
    }

    private void loadApp() {
        web.loadUrl(server() + "/pda");
    }

    /** Ha nem erheto el a szerver: hibauzenet, Ujra gomb, es a szerver cime atirhato. */
    private void showOfflinePage(String reason) {
        String html = "<html><head><meta name='viewport' content='width=device-width, initial-scale=1'></head>"
                + "<body style='background:#121212;color:#fff;font-family:sans-serif;padding:16px'>"
                + "<h2>&#x26A0;&#xFE0F; Nincs kapcsolat a szerverrel</h2>"
                + "<p style='color:#aaa'>" + escape(reason) + "</p>"
                + "<p>Ellenorizd a Wi-Fi / mobilnet kapcsolatot, majd nyomd meg az Ujra gombot.</p>"
                + "<button onclick='RaktarApp.reload()' style='width:100%;padding:16px;font-size:18px;"
                + "background:#00e676;border:none;border-radius:8px;font-weight:bold'>&#x1F504; Ujra</button>"
                + "<p style='margin-top:30px;color:#aaa'>Szerver cime:</p>"
                + "<input id='srv' value='" + escape(server()) + "' style='width:100%;padding:10px;font-size:16px'>"
                + "<button onclick='RaktarApp.setServer(document.getElementById(\"srv\").value)' "
                + "style='width:100%;margin-top:8px;padding:12px;font-size:16px'>Mentes es ujra</button>"
                + "</body></html>";
        web.loadDataWithBaseURL(null, html, "text/html", "utf-8", null);
    }

    private static String escape(String s) {
        return s == null ? "" : s.replace("&", "&amp;").replace("<", "&lt;").replace("'", "&#39;");
    }

    // ---------------- Kamera + szovegfelismeres (kezzel irt raklapszam) ----------------
    /** Elinditja a kamerat; a fenykep utan a felismert szoveget a weboldal window.onOcr(...) fuggvenye kapja meg. */
    private void startOcrPhoto() {
        try {
            File dir = new File(getCacheDir(), "ocr");
            dir.mkdirs();
            ocrFile = new File(dir, "foto.jpg");
            Uri uri = FileProvider.getUriForFile(this, getPackageName() + ".fileprovider", ocrFile);
            Intent i = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
            i.putExtra(MediaStore.EXTRA_OUTPUT, uri);
            i.addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION | Intent.FLAG_GRANT_READ_URI_PERMISSION);
            startActivityForResult(i, REQ_OCR_PHOTO);
        } catch (Exception e) {
            sendOcr(false, "", "A kamera nem indithato: " + e.getMessage());
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode != REQ_OCR_PHOTO) return;
        if (resultCode != RESULT_OK || ocrFile == null || !ocrFile.exists()) {
            sendOcr(false, "", "Nem keszult fenykep");
            return;
        }
        try {
            InputImage img = InputImage.fromFilePath(this, Uri.fromFile(ocrFile));
            TextRecognition.getClient(TextRecognizerOptions.DEFAULT_OPTIONS).process(img)
                    .addOnSuccessListener(t -> sendOcr(true, t.getText(), ""))
                    .addOnFailureListener(e -> sendOcr(false, "", "Szovegfelismeresi hiba: " + e.getMessage()));
        } catch (Exception e) {
            sendOcr(false, "", "A fenykep nem olvashato: " + e.getMessage());
        }
    }

    private void sendOcr(boolean ok, String text, String error) {
        if (web == null) return;
        try {
            JSONObject o = new JSONObject();
            o.put("ok", ok);
            o.put("text", text == null ? "" : text);
            o.put("error", error == null ? "" : error);
            final String js = "window.onOcr && window.onOcr(" + o + ")";
            runOnUiThread(() -> web.evaluateJavascript(js, null));
        } catch (Exception ignored) {
        }
    }

    class Bridge {
        /** Fenykep + szovegfelismeres (kezzel irt raklapszam a papiron). */
        @JavascriptInterface
        public void ocr() {
            runOnUiThread(MainActivity.this::startOcrPhoto);
        }

        @JavascriptInterface
        public boolean hasOcr() {
            return true;
        }

        /** A szerver jelszava (token) a PDA-n marad, nem a kodban. */
        @JavascriptInterface
        public String getToken() {
            return prefs.getString("token", "");
        }

        @JavascriptInterface
        public void setToken(String token) {
            prefs.edit().putString("token", token == null ? "" : token.trim()).apply();
        }

        @JavascriptInterface
        public void reload() {
            runOnUiThread(MainActivity.this::loadApp);
        }

        @JavascriptInterface
        public void setServer(String url) {
            String u = url == null ? "" : url.trim();
            while (u.endsWith("/")) u = u.substring(0, u.length() - 1);
            if (u.endsWith("/pda")) u = u.substring(0, u.length() - 4);
            if (!u.isEmpty()) prefs.edit().putString("server", u).apply();
            runOnUiThread(MainActivity.this::loadApp);
        }
    }

    // ---------------- Zebra DataWedge ----------------
    /** Letrehoz / frissit egy "RaktarPDA" DataWedge profilt ehhez az apphoz:
     *  olvaso bekapcsolva, a kod intentkent (broadcast) jon, billentyuzet-kimenet kikapcsolva. */
    private void configureDataWedge() {
        for (String mode : new String[]{"CREATE_IF_NOT_EXIST", "UPDATE"}) {
            Bundle profile = new Bundle();
            profile.putString("PROFILE_NAME", DW_PROFILE);
            profile.putString("PROFILE_ENABLED", "true");
            profile.putString("CONFIG_MODE", mode);

            Bundle app = new Bundle();
            app.putString("PACKAGE_NAME", getPackageName());
            app.putStringArray("ACTIVITY_LIST", new String[]{"*"});
            profile.putParcelableArray("APP_LIST", new Bundle[]{app});

            Bundle barcodeParams = new Bundle();
            barcodeParams.putString("scanner_selection", "auto");
            barcodeParams.putString("scanner_input_enabled", "true");
            Bundle barcode = new Bundle();
            barcode.putString("PLUGIN_NAME", "BARCODE");
            barcode.putString("RESET_CONFIG", "true");
            barcode.putBundle("PARAM_LIST", barcodeParams);

            Bundle intentParams = new Bundle();
            intentParams.putString("intent_output_enabled", "true");
            intentParams.putString("intent_action", SCAN_ACTION);
            intentParams.putString("intent_category", "android.intent.category.DEFAULT");
            intentParams.putString("intent_delivery", "2");   // 2 = broadcast
            Bundle intentPlugin = new Bundle();
            intentPlugin.putString("PLUGIN_NAME", "INTENT");
            intentPlugin.putString("RESET_CONFIG", "true");
            intentPlugin.putBundle("PARAM_LIST", intentParams);

            Bundle keyParams = new Bundle();
            keyParams.putString("keystroke_output_enabled", "false");
            Bundle keystroke = new Bundle();
            keystroke.putString("PLUGIN_NAME", "KEYSTROKE");
            keystroke.putString("RESET_CONFIG", "true");
            keystroke.putBundle("PARAM_LIST", keyParams);

            ArrayList<Bundle> plugins = new ArrayList<>();
            plugins.add(barcode);
            plugins.add(intentPlugin);
            plugins.add(keystroke);
            profile.putParcelableArrayList("PLUGIN_CONFIG", plugins);

            Intent i = new Intent();
            i.setAction("com.symbol.datawedge.api.ACTION");
            i.putExtra("com.symbol.datawedge.api.SET_CONFIG", profile);
            sendBroadcast(i);
        }
    }

    @Override
    protected void onResume() {
        super.onResume();
        IntentFilter filter = new IntentFilter(SCAN_ACTION);
        filter.addCategory(Intent.CATEGORY_DEFAULT);
        if (Build.VERSION.SDK_INT >= 33) {
            registerReceiver(scanReceiver, filter, Context.RECEIVER_EXPORTED);
        } else {
            registerReceiver(scanReceiver, filter);
        }
    }

    @Override
    protected void onPause() {
        super.onPause();
        try {
            unregisterReceiver(scanReceiver);
        } catch (IllegalArgumentException ignored) {
        }
    }

    @Override
    public boolean onKeyDown(int keyCode, KeyEvent event) {
        // Vissza gomb: a weboldal donti el (pl. az eladas ablak bezarasa); ha nem kezelte, kilep
        if (keyCode == KeyEvent.KEYCODE_BACK) {
            web.evaluateJavascript("window.onBack ? window.onBack() : false", value -> {
                if (!"true".equals(value)) finish();
            });
            return true;
        }
        return super.onKeyDown(keyCode, event);
    }
}
