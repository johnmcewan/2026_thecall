```markdown
# Design System Document: The Sunday Afternoon Broadcast

## 1. Overview & Creative North Star: "The Sunday Afternoon Broadcast"
This design system moves away from the sterile, high-tech minimalism of the modern web and embraces the soulful, tactile nostalgia of 1970s sports media. The Creative North Star, **"The Sunday Afternoon Broadcast,"** evokes the feeling of a faded playoff program, a grainy television signal, and the warmth of a velour tracksuit.

To break the "template" look, we reject rigid, thin-lined grids. Instead, we utilize **intentional asymmetry, heavy editorial weighting, and overlapping physical layers.** The interface should feel like a curated scrapbook or a premium sports yearbook—high-end, intentional, and deeply analog. We celebrate "ink-bleed" typography and "paper-on-paper" layering to create a digital experience that feels like it has a pulse.

---

## 2. Colors & Surface Logic
The palette is a sophisticated reimagining of 70s earth tones—saturated but "faded" by time.

### The "No-Line" Rule
**Explicit Instruction:** Designers are prohibited from using 1px solid borders for sectioning. Boundaries must be defined solely through background color shifts or subtle tonal transitions. For example, a `surface-container-low` (#f5f5dc) card sitting on a `surface` (#fbfbe2) background provides all the definition required.

### Surface Hierarchy & Nesting
Treat the UI as a series of stacked, physical sheets of heavy-stock paper.
*   **Base Layer:** `surface` (#fbfbe2) - The primary canvas.
*   **Content Wells:** `surface-container-low` (#f5f5dc) - For secondary information.
*   **Featured Elements:** `surface-container-highest` (#e4e4cc) - For high-contrast nesting.
By nesting these tiers, we create depth without the "cheapness" of digital shadows.

### The "Glass & Grain" Rule
To elevate the "faded" aesthetic, use `surface` colors with a 85-90% opacity and a `20px backdrop-blur`. This creates a "frosted acrylic" effect that feels like a vintage broadcast booth window. Always overlay a subtle, low-opacity grain texture (paper or film noise) across the entire UI to kill the "flat" digital sheen.

### Signature Gradients
CTAs and Hero elements should utilize a subtle "Sun-Drenched" gradient:
*   **Primary Action:** A linear gradient from `primary` (#9c3f00) to `primary-container` (#c45100) at a 45-degree angle. This mimics the natural fading of a vintage jersey.

---

## 3. Typography: The Editorial Voice
Our typography is a conversation between bold, nostalgic personality and modern legibility.

*   **Display & Headline (Newsreader):** This chunky, high-contrast serif provides the "soul." Use `display-lg` (3.5rem) for hero statements. It should feel massive, confident, and slightly "ink-heavy."
*   **Title & Body (Work Sans):** The clean sans-serif acts as the functional "commentary." It provides a necessary modern counterbalance to the serif’s weight.
*   **Label (Work Sans):** Used for metadata and technical specs. Always use `label-md` or `label-sm` in all-caps with a 5% letter-spacing to mimic vintage typewriter or label-maker aesthetics.

---

## 4. Elevation & Depth
In this system, depth is "baked-in" through color, not "added-on" through effects.

*   **Tonal Layering:** Avoid the "Z-axis" obsession of Material Design. Depth is achieved by placing `surface-container-lowest` (#ffffff) cards on top of `surface-dim` (#dbdcc3) backgrounds. It’s a "pressed paper" look.
*   **Ambient Shadows:** If an element must float (like a FAB or a Menu), use a "Sunlight Shadow":
    *   **Color:** `on-surface` (#1b1d0e) at 6% opacity.
    *   **Blur:** 32px.
    *   **Offset:** Y=12px (to mimic overhead afternoon sun).
*   **The Ghost Border:** If accessibility requires a border, use `outline-variant` (#dbc2b0) at **15% opacity**. Never use a 100% opaque stroke.

---

## 5. Components

### Buttons
*   **Primary:** Filled with `primary` (#9c3f00), text in `on-primary` (#ffffff). Shape: `md` (0.75rem) roundedness. 
*   **Secondary:** Filled with `secondary-container` (#fdc425), text in `on-secondary-container` (#6d5200). 
*   **Interactions:** On hover, buttons should shift slightly (2px) upwards, and the shadow should increase in blur—mimicking a physical button being released.

### Cards & Lists
*   **The Rule of Separation:** Forbid divider lines. Separate list items using `1.5rem` of vertical whitespace or by alternating background tiers (e.g., `surface` to `surface-container-low`).
*   **Rounding:** Use `xl` (1.5rem) for large containers and `md` (0.75rem) for smaller nested elements to maintain a "soft-edge" 70s feel.

### Rainbow Stripe Accents
A signature component of this system. Use a 4px tall horizontal bar consisting of four 1px stripes: `primary`, `secondary`, `tertiary`, and `outline`. Place these at the top of cards or under `headline-sm` elements to anchor the "70s Sports" theme.

### Inputs & Text Fields
*   **Style:** Minimalist but weighted. Use `surface-container-highest` for the field background with no border. 
*   **Focus State:** A 2px "Halo" in `secondary` (#785a00) at 40% opacity.

---

## 6. Do's and Don'ts

### Do:
*   **Embrace Whitespace:** Use large gaps (the "Sunday Afternoon" pace) to let typography breathe.
*   **Use Intentional Asymmetry:** Align text to the left but offset images or decorative stripes to the right to create a "layout" feel rather than a "grid" feel.
*   **Texture Over Flatness:** Apply a subtle paper grain to all `surface` containers.

### Don't:
*   **Don't use pure black (#000000):** Use `on-surface` (#1b1d0e) for all "black" needs to keep the vintage "faded" look.
*   **Don't use sharp corners:** Avoid `none` (0px) roundedness unless it's for a full-bleed background image. 70s design was "soft."
*   **Don't use 1px dividers:** If you feel the need for a line, use a 4px `surface-variant` block or a "Ghost Border" instead.
*   **Don't over-animate:** Transitions should be slow and "heavy," like a physical slide projector or a slow-motion replay. Avoid "bouncy" or "snappy" easing. Use `cubic-bezier(0.4, 0, 0.2, 1)`.```