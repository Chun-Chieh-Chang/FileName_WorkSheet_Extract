# Concept: Premium UI Standards (設計總監美學規範)

This concept defines the layout, color scheme, typography, and animation tokens that must be followed when building or refactoring web interfaces in this project. It is aimed at achieving an **"Approachable Luxury"** aesthetic.

## 🎨 Color Master Palette

We avoid high-saturation primary colors. Instead, we use Morandi-style tones, soft gradients, and high-end grays to reduce cognitive load and enhance reading clarity.

### Light Mode
*   **Base Background**: `#F9FAFB` (Cool Gray 50)
*   **Surface**: `#FFFFFF` (Pure White)
*   **Primary Text**: `#111827` (Gray 900)
*   **Secondary Text**: `#6B7280` (Gray 500)
*   **Accent/Brand**: `#3B82F6` (Royal Blue)
*   **Success**: `#10B981` (Emerald)
*   **Warning/Error**: `#EF4444` (Red)
*   **Border/Divider**: `#E5E7EB` (Gray 200)

### Dark Mode
*   **Base Background**: `#0F172A` (Slate 900)
*   **Surface**: `#1E293B` (Slate 800)
*   **Primary Text**: `#F1F5F9` (Slate 100)
*   **Secondary Text**: `#94A3B8` (Slate 400)
*   **Accent/Brand**: `#60A5FA` (Sky Blue)
*   **Success**: `#34D399` (Emerald Light)
*   **Warning/Error**: `#F87171` (Red Light)
*   **Border/Divider**: `#334155` (Slate 700)

## 📐 Layout Rules (4px Grid System)
*   All component paddings, margins, flex gaps, and grid gaps must be multiples of `4px` (specifically: 4, 8, 12, 16, 20, 24, 32, 48, 64px).
*   Avoid arbitrary pixel values like `15px`, `19px`, or `27px` to preserve structural rhythm.

## 🔤 Typography & Contrast
*   **Font Size**: Minimum `14px` for interactive UI elements.
*   **Font Family**: Use modern sans-serif fonts (e.g., Inter, Outfit, Roboto).
*   **Contrast**: Text-to-background contrast ratio must satisfy WCAG AA standards (>= 4.5:1).
*   **Line Height**: Always set line-height to `1.5` for block texts to ensure breathing room.

## ✨ Glassmorphism (毛玻璃) System
When creating floating modals, navigation bars, or glass-look cards:
*   Apply `backdrop-filter: blur(12px)`.
*   Set container transparency: `background: rgba(255, 255, 255, 0.4)` (Light) or `background: rgba(15, 23, 42, 0.4)` (Dark).
*   Enforce a subtle border: `1px solid rgba(255, 255, 255, 0.1)`.
