# Studio F Gutachten – Mehrbenutzer TEST

Separate Testanwendung. Keine produktiven Akten verwenden.

Dieses Paket basiert auf Commit 350a74d9755765d5c83ebd50797ebf92a34611f8 des Testzweigs feature/shared-storage-test-20261004. Nur index.html und shared-storage.js werden für die Laufzeit benötigt. Produktionsflows sind in dieser Version gesperrt.

## Bereitstellung

Ein neues, separates öffentliches Repository im Konto gugg-dev anlegen, beispielsweise Gutachten-Protokoll-Mehrbenutzer-TEST, sofern dieser Name verfügbar ist. Die Dateien dieses Pakets in dessen Standardbranch hochladen. Das bestehende Repository Gutachten-Protokoll-Neu nicht verändern.

GitHub Pages aus dem Standardbranch und dessen Root bereitstellen. Erst die tatsächlich bestätigte HTTPS-Adresse verwenden. In der separaten Microsoft-Testregistrierung (Client-ID 86f20e39-d4ae-4a7b-bf71-588de01d2e16) eine Single-page application-Plattform mit dieser exakten Redirect-URI einrichten. Die App berechnet sie aus location.origin + location.pathname. Mit oder ohne index.html sind unterschiedliche Rücksprungadressen: durchgängig dieselbe bestätigte URL verwenden.

Kein Client-Secret anlegen. Die eingebetteten Microsoft-IDs sind öffentliche Konfiguration, keine Zugangsdaten. Akten und Dokumente liegen in der geschützten SharePoint-Testsite und werden nicht in diesem Repository gespeichert.

## Abnahme

Mit Florian anmelden, Schema-Prüfung abwarten, synthetischen Akt anlegen. Danach mit info@studio-f.at in einem getrennten Browserkonto denselben Akt abrufen und wechselseitige Änderungen prüfen. Konflikte, lokale Entwürfe, große Protokolle sowie Ordner und Fotos gesondert testen. Kein Live-Test bisher abgeschlossen.
