# @johnsoryna/mybargraph

Eine einfache, animierte React-Komponente `HorizontalBarGraph`, die einen Wert als horizontalen Balken relativ zu einem Maximalwert anzeigt.
Entstanden im Modul 324 DevOps (Thema: Module / Komponenten / Artefakte).

## Installation

```bash
npm install @johnsoryna/mybargraph
```

`react` und `react-dom` (>= 18) sind Peer-Dependencies und müssen im Zielprojekt vorhanden sein.

## Verwendung

```jsx
import HorizontalBarGraph from '@johnsoryna/mybargraph'

function App() {
  return <HorizontalBarGraph value={30} maxvalue={100} barwidthpx={400} label="Fortschritt" />
}
```

Das CSS wird automatisch mitgeladen, es ist kein separater Import nötig.

## Props

| Prop         | Typ     | Default     | Beschreibung                                   |
|--------------|---------|-------------|------------------------------------------------|
| `value`      | number  | (required)  | Aktueller Wert, wird auf 0..maxvalue begrenzt  |
| `maxvalue`   | number  | `100`       | Maximalwert (entspricht 100 %)                 |
| `barwidthpx` | number  | `300`       | Gesamtbreite des Balkens in Pixel              |
| `label`      | string  | `''`        | Beschriftung über dem Balken                   |
| `color`      | string  | `'#2563eb'` | Farbe des gefüllten Balkens (CSS-Farbe)        |
| `showValue`  | boolean | `true`      | Wert und Prozentzahl rechts anzeigen           |

## Entwicklung

```bash
npm install
npm run demo    # Demo unter http://localhost:5173
npm run build   # Library nach dist/ bauen (ES + CJS)
```

Publizieren (Version in `package.json` vorher erhöhen):

```bash
npm version patch
npm publish --access public
```

## Repository

https://github.com/Johnsoryna/m324-react-bargraph
