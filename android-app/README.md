# Strawberry Bane — Capacitor Android shell

Packages the parent browser-game folder as an Android WebView app (`docs/android-plan.md`).

## Config

| Key | Value |
|-----|--------|
| `appId` | `de.ninea1.strawberrybane` |
| `webDir` | `www` (filled by `npm run sync:www` from parent game) |
| Plugins | `@capacitor/core`, `@capacitor/android`, `@capacitor/app` |

Web game already ships: on-screen ← → ↑, first-gesture audio unlock, Capacitor `backButton` listener when the App plugin is present.

## Workflow

```bash
cd android-app
npm install
npm run sync:www          # copy ../ game → www/
npx cap sync              # or: npm run cap:sync
npx cap open android      # needs Android Studio
```

## Environment notes

- Native `android/` project was generated with Capacitor 6 in this repo.
- **APK build / emulator / Play signing** need a local Android SDK (SDK 34+, Android Studio). This cloud agent had **no `ANDROID_HOME`** — Gradle assemble was not run here.
- Truncated `android-plan.md`: splash/adaptive icons and full Play checklist still deferred.

Portrait, safe-area CSS, and relative asset URLs are handled in the web game.
