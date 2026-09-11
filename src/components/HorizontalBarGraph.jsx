import PropTypes from 'prop-types'
import './HorizontalBarGraph.css'

/**
 * HorizontalBarGraph
 * Zeigt einen Wert (value) als horizontalen Balken relativ zu maxvalue an.
 * Der Balken wächst animiert, wenn sich value ändert.
 */
export default function HorizontalBarGraph({
  value,
  maxvalue = 100,
  barwidthpx = 300,
  label = '',
  color = '#2563eb',
  showValue = true,
}) {
  const safeMax = maxvalue > 0 ? maxvalue : 1
  const clamped = Math.min(Math.max(value, 0), safeMax)
  const percent = Math.round((clamped / safeMax) * 100)

  return (
    <div className="hbg" style={{ width: `${barwidthpx}px` }}>
      {(label || showValue) && (
        <div className="hbg__header">
          <span className="hbg__label">{label}</span>
          {showValue && (
            <span className="hbg__value">
              {clamped} / {safeMax} ({percent}%)
            </span>
          )}
        </div>
      )}
      <div
        className="hbg__track"
        role="progressbar"
        aria-valuemin={0}
        aria-valuemax={safeMax}
        aria-valuenow={clamped}
        aria-label={label || 'bar graph'}
      >
        <div
          className="hbg__bar"
          style={{ width: `${percent}%`, backgroundColor: color }}
        />
      </div>
    </div>
  )
}

HorizontalBarGraph.propTypes = {
  /** Aktueller Wert, wird auf 0..maxvalue begrenzt */
  value: PropTypes.number.isRequired,
  /** Maximalwert (100 %) */
  maxvalue: PropTypes.number,
  /** Gesamtbreite des Balkens in Pixel */
  barwidthpx: PropTypes.number,
  /** Beschriftung über dem Balken */
  label: PropTypes.string,
  /** Farbe des gefüllten Balkens (CSS-Farbe) */
  color: PropTypes.string,
  /** Wert und Prozent rechts anzeigen */
  showValue: PropTypes.bool,
}
