import { LEGEND } from '../status'

export function Legend() {
  return (
    <ul className="legend" aria-label="Status legend">
      {LEGEND.map(({ key, label }) => (
        <li key={key} className="legend-item">
          <span className={`swatch status-${key}`} aria-hidden="true">
            {key}
          </span>
          {label}
        </li>
      ))}
    </ul>
  )
}
