# Design System Strategy: The Luminous Analyst

## 1. Overview & Creative North Star
The Creative North Star for this design system is **"The Luminous Analyst."** 

In a world of flat, uninspired SaaS dashboards, this system moves toward a high-end editorial aesthetic that mimics a sophisticated command center. We are moving away from the "grid of boxes" and toward a multi-layered, immersive environment. The experience is defined by **Tonal Depth** and **Vibrant Data Luminescence**. By utilizing a palette of deep charcoals and intentional background shifts, we create a canvas where data doesn't just sit on the screen—it glows with purpose. 

The layout breaks the standard "template" look through:
*   **Intentional Asymmetry:** Grouping global filters in glassmorphic sidebars while allowing data visualizations to breathe in expansive, high-contrast sections.
*   **Overlapping Elevations:** Using surface tiers to stack elements, making the UI feel like physical layers of high-tech instrumentation.
*   **Editorial Scale:** A dramatic contrast between large, characterful headlines and ultra-clean, high-density functional data.

---

## 2. Colors & Surface Philosophy
The palette is anchored in `#070d1f` (Background/Surface), creating a "void" that allows primary accents to pop with maximum chroma.

*   **Primary (`#3bbffa`):** Used for interactive focus and "success" momentum in data.
*   **Secondary & Tertiary (`#69f6b8`, `#ffb148`):** Reserved for high-priority data differentiation.
*   **The "No-Line" Rule:** We strictly prohibit 1px solid borders for structural sectioning. To separate a sidebar from a main feed, use a shift from `surface` to `surface-container-low`. Boundaries are felt, not seen.
*   **Surface Hierarchy & Nesting:**
    *   **Level 0 (Base):** `surface` (#070d1f) for the global background.
    *   **Level 1 (Sections):** `surface-container` (#11192e) for primary content areas.
    *   **Level 2 (Cards):** `surface-container-high` (#171f36) or `surface-container-highest` (#1c253e) for data modules.
*   **The "Glass & Gradient" Rule:** Floating panels (like tooltips or global filters) must use Glassmorphism. Apply `surface-variant` with a 60% opacity and a `backdrop-filter: blur(20px)`. Main CTAs should utilize a subtle linear gradient from `primary` to `primary-dim` to provide a "machined" premium finish.

---

## 3. Typography: The Editorial Balance
We utilize two distinct typefaces to balance authority with technical precision.

*   **Display & Headlines (Manrope):** A geometric sans-serif with high personality. Use `display-lg` (3.5rem) for hero metrics to establish an editorial "magazine" feel. Manrope’s wide stance conveys modern sophistication.
*   **Body & UI (Inter):** The workhorse for legibility. Inter is used for all functional data, tables, and labels. 
*   **Hierarchy as Identity:** By pairing a large `headline-md` (1.75rem) in Manrope with a tiny, uppercase `label-sm` (0.6875rem) in Inter, we create a "professional analysis" aesthetic found in high-end financial reports.

---

## 4. Elevation & Depth
Depth is the primary navigator in this system. We avoid shadows that look like "drops"; we want "glows" and "lifts."

*   **The Layering Principle:** Place `surface-container-lowest` cards on `surface-container-low` sections. This creates "negative depth," making the card feel like a carved-out pocket of information.
*   **Ambient Shadows:** For floating elements, use a diffused shadow: `box-shadow: 0 20px 40px rgba(0, 0, 0, 0.4)`. The shadow must feel like it belongs to the dark environment, never grey or muddy.
*   **The "Ghost Border" Fallback:** If a border is required for high-density tables, use the `outline-variant` token at **15% opacity**. This creates a "hairline" effect that is visible but does not break the visual flow.
*   **Luminescent Charts:** Charts must feature a 5px - 10px outer glow (drop-shadow) using the data's specific color (e.g., `primary` or `secondary`) at 30% opacity to mimic a glowing CRT monitor.

---

## 5. Components

### Buttons
*   **Primary:** `primary` background, `on-primary` text. `xl` (0.75rem) roundedness. No border.
*   **Secondary:** `outline` ghost border (20% opacity) with `on-surface` text.
*   **States:** On hover, use `surface-tint` to create a subtle "inner glow" effect rather than a simple color change.

### Data Tables (High-Density)
*   **Forbid Dividers:** Use `surface-container-low` vs. `surface-container-high` row stripping for separation.
*   **Header:** `label-md` in `on-surface-variant`, all-caps, with 1.5 spacing.
*   **Padding:** Use `spacing.3` (0.6rem) for vertical cell padding to maximize data density without sacrificing "breathability."

### Cards
*   **Style:** No borders. Use `surface-container-highest` for the background. 
*   **Glow:** Apply a subtle `primary` top-border (2px) only on "Active" or "Featured" cards to draw the eye.

### Input Fields
*   **Surface:** `surface-container-lowest` (#000000) to create a "recessed" look.
*   **Active State:** A 1px `primary` border with a 4px `primary` outer glow (10% opacity).

---

## 6. Do's and Don'ts

### Do:
*   **Do** use vertical white space (from the 8 or 10 spacing scale) to group content instead of lines.
*   **Do** use vibrant gradients on sparklines to show trend momentum.
*   **Do** leverage `surface-bright` (#222b47) for tiny UI accents like notification pips or active tab indicators.

### Don't:
*   **Don't** use 100% white (#FFFFFF). Always use `on-surface` (#dfe4fe) to prevent eye strain in dark mode.
*   **Don't** use standard 1px borders to box in your charts; let the data breath against the container background.
*   **Don't** use generic drop shadows. If it doesn't look like an ambient light source, it doesn't belong in this system.