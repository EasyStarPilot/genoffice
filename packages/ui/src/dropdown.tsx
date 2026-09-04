/**
 * Shared select replacement: a framed trigger (current option + caret) opening
 * a themed, anchored popover of options — native <select> popups are OS-drawn
 * and ignore the app theme. Options render arbitrary content (line previews,
 * text) and stay keyboard-operable (arrows / Enter / Escape / Home / End while
 * the trigger keeps focus). The popover is position: fixed + CSS-anchored to
 * the trigger so scroll-container overflow never clips it, with a flip-block
 * fallback near the viewport bottom.
 *
 * `searchable` adds a filter box atop the popover for long or dynamically
 * fetched lists (e.g. OpenRouter's full model catalog) — typing narrows
 * options by label substring; arrow keys and Enter keep working from there.
 *
 * Styling comes from dropdown.css (gs-dd* classes, token colors only); apps
 * size the control via `className` on the wrapper.
 */
import React, { useEffect, useMemo, useRef, useState } from 'react'
import { useDismissablePopover } from './popover-dismiss'

export interface DropdownOption<K extends string = string> {
  readonly value: K
  /** Accessible name (aria-label/title); also the visible text when `render` is omitted. */
  readonly label: string
  readonly render?: React.ReactNode
  /** Shown but not pickable (placeholder rows like "Auto", unavailable modes). */
  readonly disabled?: boolean
}

export function Dropdown<K extends string>({
  value,
  options,
  onPick,
  className,
  ariaLabel,
  disabled,
  tip,
  ariaRequired,
  ariaInvalid,
  searchable,
  filterPlaceholder,
  emptyText,
}: {
  readonly value: K
  readonly options: ReadonlyArray<DropdownOption<K>>
  readonly onPick: (value: K) => void
  /** Extra class on the wrapper (apps set width there). */
  readonly className?: string
  /** Control name for screen readers; defaults to the current option's label. */
  readonly ariaLabel?: string
  readonly disabled?: boolean
  /** ScreenTip text (data-tip on the trigger). */
  readonly tip?: string
  /** Form semantics passthrough (AcroForm widgets etc.). */
  readonly ariaRequired?: boolean
  readonly ariaInvalid?: boolean
  /** Renders a filter box atop the popover, narrowing options by label substring. */
  readonly searchable?: boolean
  /** Placeholder for the filter box (only used when searchable). */
  readonly filterPlaceholder?: string
  /** Shown in place of the list when a search matches nothing (only used when searchable). */
  readonly emptyText?: string
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const [active, setActive] = useState(0)
  const [query, setQuery] = useState('')
  const popRef = useRef<HTMLDivElement>(null)
  const wrapRef = useRef<HTMLSpanElement>(null)
  const btnRef = useRef<HTMLButtonElement>(null)
  const searchRef = useRef<HTMLInputElement>(null)
  // guarded (capture-phase) dismissal: a press on another dropdown's trigger
  // must close this one even though that trigger stops mousedown propagation
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })

  const visible = useMemo(() => {
    if (!searchable || !query.trim()) return options
    const q = query.trim().toLowerCase()
    return options.filter((o) => o.label.toLowerCase().includes(q))
  }, [options, searchable, query])
  // clamped at read time (not stored) so a filter that shrinks the list can't
  // leave `active` pointing past its end
  const activeIdx = Math.min(active, Math.max(0, visible.length - 1))

  useEffect(() => {
    if (!open) return
    popRef.current?.querySelectorAll('.gs-dd-item')[activeIdx]?.scrollIntoView?.({
      block: 'nearest',
    })
  }, [open, activeIdx])

  useEffect(() => {
    if (open && searchable) searchRef.current?.focus()
  }, [open, searchable])

  // No fallback to options[0]: an off-list value (e.g. a document-only font)
  // must read as itself, not masquerade as the first option
  const current = options.find((o) => o.value === value)
  const openList = () => {
    const i = options.findIndex((o) => o.value === value)
    setActive(i < 0 ? 0 : i)
    setQuery('')
    setOpen(true)
  }
  const closeList = () => {
    setOpen(false)
    setQuery('')
  }
  const pick = (o: DropdownOption<K>) => {
    if (o.disabled) return
    closeList()
    onPick(o.value)
  }
  /** Arrow/Home/End/Enter/Escape, shared by the trigger and the search box. */
  const navigate = (e: React.KeyboardEvent, pickOnSpace: boolean) => {
    if (e.key === 'Escape') closeList()
    else if (e.key === 'ArrowDown') setActive((i) => Math.min(visible.length - 1, i + 1))
    else if (e.key === 'ArrowUp') setActive((i) => Math.max(0, i - 1))
    else if (e.key === 'Home') setActive(0)
    else if (e.key === 'End') setActive(visible.length - 1)
    else if (e.key === 'Enter' || (pickOnSpace && e.key === ' ')) {
      const o = visible[activeIdx]
      if (o) pick(o)
    } else return
    e.preventDefault()
    // handled keys stay ours while the list is open: a bubbling Escape would
    // close the hosting modal, bubbling arrows would nudge canvas elements
    e.stopPropagation()
  }
  const onKeyDown = (e: React.KeyboardEvent) => {
    if (!open) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter' || e.key === ' ') {
        e.preventDefault()
        openList()
      }
      return
    }
    navigate(e, true)
  }
  // space must type into the filter box, not act as a pick shortcut
  const onSearchKeyDown = (e: React.KeyboardEvent) => {
    navigate(e, false)
    if (e.key === 'Escape') btnRef.current?.focus()
  }
  return (
    <span ref={wrapRef} className={`gs-dd${className ? ` ${className}` : ''}`}>
      <button
        ref={btnRef}
        type="button"
        className="gs-dd-btn"
        disabled={disabled}
        data-value={value}
        data-tip={tip}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label={ariaLabel ?? current?.label ?? value}
        aria-required={ariaRequired}
        aria-invalid={ariaInvalid}
        onClick={() => (open ? closeList() : openList())}
        onKeyDown={onKeyDown}
        onBlur={(e) => {
          // native selects close on focus loss (Tab); staying inside the wrapper
          // (clicking an option, or the search box taking focus) must not dismiss
          if (!wrapRef.current?.contains(e.relatedTarget as Node)) closeList()
        }}
      >
        <span className="gs-dd-value">{current ? (current.render ?? current.label) : value}</span>
        <span className="gs-dd-caret" aria-hidden="true">
          <svg width="11" height="11" viewBox="0 0 24 24" fill="none">
            <path
              d="M5.5 9.25 12 15.75l6.5-6.5"
              stroke="currentColor"
              strokeWidth="2.6"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </span>
      </button>
      {open && (
        <div ref={popRef} className="gs-dd-pop" role="listbox">
          {searchable && (
            <div className="gs-dd-search-wrap">
              <input
                ref={searchRef}
                type="text"
                className="gs-dd-search"
                value={query}
                placeholder={filterPlaceholder}
                spellCheck={false}
                autoComplete="off"
                onChange={(e) => {
                  setQuery(e.target.value)
                  setActive(0)
                }}
                onKeyDown={onSearchKeyDown}
                onBlur={(e) => {
                  if (!wrapRef.current?.contains(e.relatedTarget as Node)) closeList()
                }}
              />
            </div>
          )}
          {visible.length === 0 ? (
            <div className="gs-dd-empty">{emptyText ?? 'No matches'}</div>
          ) : (
            visible.map((o, i) => (
              <button
                key={o.value}
                type="button"
                role="option"
                // menu-button pattern: options never join the tab order (focus
                // lives on the trigger; Tab away closes via its onBlur)
                tabIndex={-1}
                disabled={o.disabled}
                aria-selected={o.value === value}
                aria-label={o.label}
                data-value={o.value}
                title={o.label}
                className={`gs-dd-item${o.value === value ? ' selected' : ''}${i === activeIdx && !o.disabled ? ' active' : ''}`}
                onMouseEnter={() => setActive(i)}
                // menu-button pattern: options never take focus, so picking one
                // can't blur focus-scoped hosts (the PDF text editor commits its
                // draft on focus leaving the edit bar)
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => pick(o)}
              >
                {o.render ?? o.label}
              </button>
            ))
          )}
        </div>
      )}
    </span>
  )
}
