/* Monday Nite Footballers — Shared Tailwind Configuration
   Load this BEFORE <script id="tailwind-config"> and the Tailwind CDN.
   Usage in each page:
       <script src="mnf-config.js"></script>
       <script id="tailwind-config">tailwind.config = window._mnfConfig</script>
       <script src="https://cdn.tailwindcss.com?plugins=forms,container-queries"></script>
*/
window._mnfConfig = {
  darkMode: "class",
  theme: {
    extend: {
      colors: {
        /* Surface hierarchy */
        "surface":                  "#fbfbe2",
        "surface-bright":          "#fbfbe2",
        "surface-dim":             "#dbdcc4",
        "surface-variant":         "#e4e4cc",
        "surface-container-lowest":"#ffffff",
        "surface-container-low":   "#f5f5dc",
        "surface-container":       "#efefd7",
        "surface-container-high":  "#eaead1",
        "surface-container-highest":"#e4e4cc",
        "surface-tint":            "#9f4103",
        /* Brand */
        "primary":                 "#762e00",
        "primary-container":       "#9c3f00",
        "primary-fixed":           "#ffdbcc",
        "primary-fixed-dim":       "#ffb693",
        "on-primary":              "#ffffff",
        "on-primary-container":    "#ffc9b0",
        "on-primary-fixed":        "#351000",
        "on-primary-fixed-variant":"#7a3000",
        "inverse-primary":         "#ffb693",
        /* Secondary */
        "secondary":               "#785a00",
        "secondary-container":     "#fdc425",
        "secondary-fixed":         "#ffdf9a",
        "secondary-fixed-dim":     "#f6be1d",
        "on-secondary":            "#ffffff",
        "on-secondary-container":  "#6d5200",
        "on-secondary-fixed":      "#251a00",
        "on-secondary-fixed-variant":"#5a4300",
        /* Tertiary */
        "tertiary":                "#004388",
        "tertiary-container":      "#005ab3",
        "tertiary-fixed":          "#d6e3ff",
        "tertiary-fixed-dim":      "#aac7ff",
        "on-tertiary":             "#ffffff",
        "on-tertiary-container":   "#c0d5ff",
        "on-tertiary-fixed":       "#001b3e",
        "on-tertiary-fixed-variant":"#00458d",
        /* On-surface / text */
        "on-surface":              "#1b1d0e",
        "on-surface-variant":      "#564239",
        "on-background":           "#1b1d0e",
        "inverse-surface":         "#303221",
        "inverse-on-surface":      "#f2f2da",
        /* Background */
        "background":              "#fbfbe2",
        /* Outline */
        "outline":                 "#8a7268",
        "outline-variant":         "#ddc1b4",
        /* Error */
        "error":                   "#ba1a1a",
        "error-container":         "#ffdad6",
        "on-error":                "#ffffff",
        "on-error-container":      "#93000a",
        /* Alias used in people.html */
        "archive-blue":            "#004388",
      },
      borderRadius: {
        DEFAULT: "0.125rem",
        lg:      "0.25rem",
        xl:      "0.5rem",
        full:    "0.75rem",
      },
      fontFamily: {
        headline: ["Newsreader", "serif"],
        body:     ["Work Sans", "sans-serif"],
        label:    ["Work Sans", "sans-serif"],
      },
    },
  },
};
