# Teal Tower Analytics 📊

A high-performance, modern data visualization dashboard built with Astro, TypeScript, and Firebase. 
Teal Tower is designed to process and analyze local Excel/CSV datasets instantly, offering seamless cloud synchronization and beautiful "Liquid Glass" UI aesthetics.

## ✨ Features

- **Blazing Fast Processing:** Client-side parsing of Excel and CSV files.
- **Liquid Glass UI:** Modern glassmorphism design with hardware-accelerated animations and premium styling.
- **Data Visualization:** Interactive charts, accumulation graphs, and distribution stats.
- **Cloud Sync:** Persistent state synchronization across devices using Firebase Firestore.
- **Fully Responsive:** Optimized for both desktop and mobile accessibility with native touch-scrolling.

## 🛠️ Tech Stack

- **Framework:** [Astro](https://astro.build/)
- **Language:** TypeScript
- **Styling:** Custom Vanilla CSS (Liquid Glass Aesthetics)
- **Database:** Firebase Firestore
- **Charts:** Chart.js
- **File Parsing:** SheetJS (xlsx)

## 🚀 Getting Started

### Prerequisites
Make sure you have Node.js version 18+ installed.

### Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Set up your environment variables. Create a `.env` file in the root directory:
   ```env
   PUBLIC_FIREBASE_API_KEY="your-api-key"
   ```
4. Start the development server:
   ```bash
   npm run dev
   ```

## 📜 License

This project is licensed under the MIT License.
