package com.lectordocs.app;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.content.ActivityNotFoundException;
import android.content.Intent;
import android.database.Cursor;
import android.net.Uri;
import android.os.Bundle;
import android.provider.OpenableColumns;
import android.webkit.JavascriptInterface;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;

import org.json.JSONObject;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * Contenedor nativo mínimo: la interfaz y los visores viven en assets/www.
 * Las páginas se sirven desde https://appassets.androidplatform.net para que
 * los módulos ES de pdf.js funcionen (no cargan desde file://).
 * El documento abierto se entrega al JS por la ruta /doc/actual, sin pasar
 * por base64, así que los archivos grandes no saturan la memoria.
 */
public class MainActivity extends Activity {

    private static final String HOST = "appassets.androidplatform.net";
    private static final String START_URL = "https://" + HOST + "/www/index.html";
    private static final int PICK_FILE = 41;

    private static final String[] MIME_TYPES = {
            "application/pdf",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
    };

    private static final Map<String, String> CONTENT_TYPES = new HashMap<>();
    static {
        CONTENT_TYPES.put("html", "text/html");
        CONTENT_TYPES.put("js", "application/javascript");
        CONTENT_TYPES.put("mjs", "application/javascript");
        CONTENT_TYPES.put("css", "text/css");
        CONTENT_TYPES.put("svg", "image/svg+xml");
        CONTENT_TYPES.put("png", "image/png");
        CONTENT_TYPES.put("json", "application/json");
        CONTENT_TYPES.put("woff2", "font/woff2");
    }

    private WebView webView;
    private volatile Uri currentUri;
    private boolean pageReady = false;
    private Uri pendingUri;

    @SuppressLint("SetJavaScriptEnabled")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        webView = new WebView(this);
        setContentView(webView);

        WebSettings s = webView.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setAllowFileAccess(false);
        s.setAllowContentAccess(false);
        s.setBuiltInZoomControls(true);
        s.setDisplayZoomControls(false);
        s.setSupportZoom(true);
        s.setUseWideViewPort(true);
        s.setLoadWithOverviewMode(true);

        webView.addJavascriptInterface(new Bridge(), "Android");
        webView.setWebViewClient(new WebViewClient() {
            @Override
            public WebResourceResponse shouldInterceptRequest(WebView view, WebResourceRequest request) {
                return serve(request.getUrl());
            }

            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                Uri u = request.getUrl();
                if (HOST.equals(u.getHost())) return false;
                // Enlaces dentro de documentos: se abren en el navegador del teléfono
                try {
                    startActivity(new Intent(Intent.ACTION_VIEW, u));
                } catch (ActivityNotFoundException ignored) { }
                return true;
            }

            @Override
            public void onPageFinished(WebView view, String url) {
                pageReady = true;
                if (pendingUri != null) {
                    Uri u = pendingUri;
                    pendingUri = null;
                    deliver(u);
                }
            }
        });

        webView.loadUrl(START_URL);
        handleIntent(getIntent());
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        setIntent(intent);
        handleIntent(intent);
    }

    private void handleIntent(Intent intent) {
        if (intent == null) return;
        Uri uri = null;
        if (Intent.ACTION_VIEW.equals(intent.getAction())) {
            uri = intent.getData();
        } else if (Intent.ACTION_SEND.equals(intent.getAction())) {
            uri = intent.getParcelableExtra(Intent.EXTRA_STREAM);
        }
        if (uri == null) return;
        if (pageReady) deliver(uri); else pendingUri = uri;
    }

    /** Avisa a la interfaz que hay un documento listo en /doc/actual. */
    private void deliver(Uri uri) {
        currentUri = uri;
        try {
            JSONObject info = new JSONObject();
            info.put("name", displayName(uri));
            info.put("mime", getContentResolver().getType(uri));
            info.put("uri", uri.toString());
            info.put("size", fileSize(uri));
            final String js = "window.onNativeFile && window.onNativeFile(" + info + ")";
            runOnUiThread(() -> webView.evaluateJavascript(js, null));
        } catch (Exception e) {
            toast("No se pudo abrir el archivo");
        }
    }

    private WebResourceResponse serve(Uri url) {
        if (!HOST.equals(url.getHost()) || url.getPath() == null) return null;
        String path = url.getPath();
        try {
            if (path.startsWith("/doc/")) {
                Uri u = currentUri;
                if (u == null) return notFound();
                InputStream in = getContentResolver().openInputStream(u);
                WebResourceResponse r = new WebResourceResponse("application/octet-stream", null, in);
                r.setResponseHeaders(noCache());
                return r;
            }
            if (path.startsWith("/www/")) {
                String asset = path.substring(1);
                if (asset.contains("..")) return notFound();
                String ext = asset.substring(asset.lastIndexOf('.') + 1).toLowerCase(Locale.ROOT);
                String type = CONTENT_TYPES.containsKey(ext) ? CONTENT_TYPES.get(ext) : "application/octet-stream";
                InputStream in = getAssets().open(asset);
                return new WebResourceResponse(type, "UTF-8", in);
            }
        } catch (Exception e) {
            return notFound();
        }
        return notFound();
    }

    private WebResourceResponse notFound() {
        WebResourceResponse r = new WebResourceResponse("text/plain", "UTF-8", null);
        r.setStatusCodeAndReasonPhrase(404, "Not Found");
        return r;
    }

    private Map<String, String> noCache() {
        Map<String, String> h = new HashMap<>();
        h.put("Cache-Control", "no-store");
        return h;
    }

    private String displayName(Uri uri) {
        String name = null;
        try (Cursor c = getContentResolver().query(uri, new String[]{OpenableColumns.DISPLAY_NAME}, null, null, null)) {
            if (c != null && c.moveToFirst()) name = c.getString(0);
        } catch (Exception ignored) { }
        if (name == null) {
            name = uri.getLastPathSegment();
            if (name != null && name.contains("/")) name = name.substring(name.lastIndexOf('/') + 1);
        }
        return name == null ? "documento" : name;
    }

    private long fileSize(Uri uri) {
        try (Cursor c = getContentResolver().query(uri, new String[]{OpenableColumns.SIZE}, null, null, null)) {
            if (c != null && c.moveToFirst() && !c.isNull(0)) return c.getLong(0);
        } catch (Exception ignored) { }
        return -1;
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode != PICK_FILE || resultCode != RESULT_OK || data == null || data.getData() == null) return;
        Uri uri = data.getData();
        // Permiso persistente para poder reabrirlo desde "Recientes"
        try {
            getContentResolver().takePersistableUriPermission(uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
        } catch (Exception ignored) { }
        deliver(uri);
    }

    @Override
    public void onBackPressed() {
        webView.evaluateJavascript("window.handleBack ? window.handleBack() : false", value -> {
            if (!"true".equals(value)) MainActivity.super.onBackPressed();
        });
    }

    private void toast(String msg) {
        runOnUiThread(() -> Toast.makeText(this, msg, Toast.LENGTH_SHORT).show());
    }

    /** Métodos que la interfaz web puede llamar como Android.xxx() */
    private class Bridge {
        @JavascriptInterface
        public void pickFile() {
            Intent i = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            i.addCategory(Intent.CATEGORY_OPENABLE);
            i.setType("*/*");
            i.putExtra(Intent.EXTRA_MIME_TYPES, MIME_TYPES);
            i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION);
            runOnUiThread(() -> {
                try {
                    startActivityForResult(i, PICK_FILE);
                } catch (ActivityNotFoundException e) {
                    toast("No hay un gestor de archivos disponible");
                }
            });
        }

        /** Reabre un documento de la lista de recientes. Devuelve false si ya no hay acceso. */
        @JavascriptInterface
        public boolean openUri(String uriString) {
            Uri uri = Uri.parse(uriString);
            try (InputStream in = getContentResolver().openInputStream(uri)) {
                if (in == null) return false;
            } catch (Exception e) {
                return false;
            }
            deliver(uri);
            return true;
        }

        @JavascriptInterface
        public void share(String uriString) {
            Uri uri = Uri.parse(uriString);
            Intent i = new Intent(Intent.ACTION_SEND);
            String type = getContentResolver().getType(uri);
            i.setType(type == null ? "*/*" : type);
            i.putExtra(Intent.EXTRA_STREAM, uri);
            i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            runOnUiThread(() -> startActivity(Intent.createChooser(i, "Compartir")));
        }

        @JavascriptInterface
        public void toast(String msg) {
            MainActivity.this.toast(msg);
        }
    }
}
