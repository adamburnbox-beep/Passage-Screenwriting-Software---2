# Antigravity Layout Directive: E-Reader Styling Constraints

You are strictly required to refactor the WPF visual elements within MainWindow.xaml and the respective theme dictionaries to enforce a flat, binary contrast E-Reader style. Follow these structural mandates precisely:

1. GEOMETRIC FLATTENING: Set 'CornerRadius' to '0' globally across all styles, templates, borders, and custom buttons. 
2. DEPTH ELIMINATION: Locate and delete all '<DropShadowEffect />' blocks and parent '<Border.Effect>' elements.
3. TRANSITION BYPASS: Change all 'PopupAnimation' values from 'Slide' or 'Fade' to 'None' globally.
4. CONTRAST LAYER SYSTEM: In LightTheme.xaml and DarkTheme.xaml, convert any low-opacity or transparent color brush variants (e.g., alpha-blended grid dividers) into solid, opaque hexadecimal tones.