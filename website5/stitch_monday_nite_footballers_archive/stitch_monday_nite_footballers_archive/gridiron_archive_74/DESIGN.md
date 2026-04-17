# Design System Strategy: The Sunday Afternoon Broadcast

## 1. Overview & Creative North Star
The Creative North Star for this design system is **"The Living Archive."** We are not building a modern dashboard; we are curating a tactile, 1970s-inspired sports research folio. This system rejects the sterile, "app-like" perfection of modern SaaS in favor of an editorial experience that feels printed, collected, and weathered. 

To break the "template" look, we utilize **Intentional Asymmetry**. Containers should not always align to a rigid center; they should feel like documents dropped onto a desk. We use **Paper-on-Paper Layering** to create depth, where the UI feels like a series of physical artifacts—stat sheets, photographs, and clippings—stacked intentionally to guide the eye.

## 2. Colors: The Earthy Palette
The color language is rooted in the faded, sun-drenched broadcast aesthetics of the 1970s. It avoids the harsh blacks and neon whites of modern digital interfaces.

### The "No-Line" Rule
**Traditional 1px solid borders are strictly prohibited.** Sectioning must be achieved through tonal shifts or physical stacking. To define a new area, shift the background color from `surface` to `surface-container-low` or `surface-container-highest`. 

### Surface Hierarchy & Nesting
Treat the screen as a physical desk.
- **Base Layer:** `surface` (#fbfbe2) – The desktop.
- **Section Layer:** `surface-container` (#efefd7) – A folder or large document.
- **Interactive Layer:** `surface-container-highest` (#e4e4cc) – A highlighted clipping or active card.
- **The "Glass & Gradient" Rule:** For floating elements or "on-air" overlays, use a semi-transparent `surface-bright` with a `20px` backdrop-blur to mimic frosted acetate overlays used in vintage broadcast graphics.

### Signature Textures
- **The Ink Bleed:** Headlines in `primary` (#762e00) should occasionally utilize a very subtle outer glow (0.5px blur) of the same color to simulate ink spreading into high-quality paper stock.
- **Rainbow Accents:** Use the `secondary` (#785a00), `tertiary` (#004388), and `error` (#ba1a1a) tokens in thin, grouped "rainbow stripes" (3px height) to denote section breaks or categories, nodding to vintage television test patterns.

## 3. Typography: Editorial Authority
Typography is our primary tool for storytelling. We pair a chunky, authoritative serif with a functional, modern sans-serif to bridge the gap between "broadcast news" and "data research."

*   **Display & Headlines (Newsreader):** This is our "Broadcaster" voice. It must be set with tighter letter-spacing (-0.02em) and high contrast. Use `display-lg` for heroic moments and `headline-md` for standard section starts.
*   **Body & Labels (Work Sans):** This is our "Statistician" voice. Work Sans provides a clean, neutral counterpoint to the Newsreader serif. 
*   **The "Ink Weight" Principle:** Headlines should always use `on-surface` (#1b1d0e) or `primary` (#762e00) to ensure they feel "stamped" onto the page.

## 4. Elevation & Depth: Tonal Layering
In this system, depth is a product of material stacking, not artificial lighting.

*   **The Layering Principle:** Place a `surface-container-lowest` (#ffffff) card on top of a `surface-container-low` (#f5f5dc) background. This creates a "lift" that feels like a fresh sheet of paper resting on an aged one.
*   **Ambient Shadows:** Use shadows sparingly. When required for "floating" cards, use a large blur (24px+) at 6% opacity, using the `on-background` (#1b1d0e) color to ensure the shadow feels like soft, ambient room light.
*   **The "Ghost Border" Fallback:** If a boundary is required for accessibility, use `outline-variant` (#ddc1b4) at 15% opacity. Never use a 100% opaque border.
*   **Intentional Asymmetry:** Offset "stacked" layers by 4px or 8px horizontally or vertically. This reinforces the "hand-placed" feel of a physical research project.

## 5. Components

### Buttons
*   **Primary:** Solid `primary_container` (#9c3f00) with `on_primary` (#ffffff) text. Use `rounded-sm` (2px) to keep edges sharp and "printed."
*   **Secondary:** No fill. Use a `surface-container-highest` background on hover. Use `title-sm` typography.
*   **Tertiary:** Underlined `Work Sans` text. The underline should be 2px thick and offset by 4px.

### Cards & Lists
*   **Constraint:** No dividers. Separate list items using 16px of vertical whitespace or by alternating backgrounds between `surface` and `surface-container-low`.
*   **The "Clipping" Card:** Use `surface-container-highest` with an intentional 2-degree rotation for "featured" research notes to break the grid.

### Chips (Tags)
*   **Selection Chips:** Use `secondary_container` (#fdc425) with `on_secondary_container` (#6d5200). These should look like highlighter marks.
*   **Styling:** Rectangular with `rounded-none`.

### Input Fields
*   **Style:** A simple underline using `outline` (#8a7268) instead of a boxed container. Labels sit in `label-md` (Work Sans) above the line.
*   **Focus State:** The underline thickens to 2px and changes to `primary` (#762e00).

### Custom Component: The "Broadcast Stripe"
A decorative element consisting of four 4px stripes using `primary`, `secondary`, `tertiary`, and `error` tokens. Use this at the top of "Data Sheets" or as a left-hand accent for active navigation items.

## 6. Do's and Don'ts

### Do
*   **Do** use Newsreader for any text that carries "editorial" weight (quotes, headers, big numbers).
*   **Do** allow elements to overlap. A photo can partially cover a headline to create a "scrapbook" feel.
*   **Do** use `surface-dim` for "archived" or "disabled" content to make it look faded by the sun.

### Don't
*   **Don't** use pure black (#000000) or pure white (#FFFFFF) except in extreme UI cases. Use the surface tokens provided.
*   **Don't** use standard "Material Design" shadows. They are too digital for this aesthetic.
*   **Don't** align everything to a perfect vertical line. Nudge secondary elements 8-16px off-grid to create visual interest.