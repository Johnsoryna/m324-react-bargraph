import React, { useState } from 'react'
import { createRoot } from 'react-dom/client'
// Direkt aus dem Quellcode importieren, nicht aus dist/, damit man beim Entwickeln sofort sieht, was man ändert.
import HorizontalBarGraph from '../../src/components/HorizontalBarGraph.jsx'

const App = () => {
  const [value, setValue] = useState(30)

  return (
    <div style={{ fontFamily: 'system-ui, sans-serif', padding: 24 }}>
      <h1>MyBargraph – Demo</h1>
      <p>
        <label>
          Wert: {value}{' '}
          <input
            type="range"
            min="0"
            max="100"
            value={value}
            onChange={(e) => setValue(Number(e.target.value))}
          />
        </label>
      </p>
      <HorizontalBarGraph value={value} maxvalue={100} barwidthpx={400} label="Fortschritt" />
      <br />
      <HorizontalBarGraph value={value} maxvalue={50} barwidthpx={250} label="Max 50" color="#16a34a" />
      <br />
      <HorizontalBarGraph value={value} maxvalue={100} barwidthpx={300} color="#dc2626" showValue={false} />
    </div>
  )
}

createRoot(document.getElementById('root')).render(<App />)
