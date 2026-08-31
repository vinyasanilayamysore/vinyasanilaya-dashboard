# Vinyasa Nilaya Dashboard

A professional property management and guest relations dashboard built on **Google Apps Script (GAS)** and **Bootstrap**. This application provides a centralized interface for tracking bookings, managing secure guest documentation, and monitoring financial performance for Vinyasa Nilaya.

## 🚀 Key Features

### 📊 Financial & Analytics Dashboard
*   **Real-time Metrics**: Instantly view "Period Revenue" and "Check-in Counts" alongside "All-time" totals.
*   **Dynamic Filtering**: Sift through data using year, month, booking source (Airbnb vs. Personal), and floor assignment filters.
*   **Unified Search**: Quickly find guest records by name or mobile number.

### 🛎️ Booking & Stay Management
*   **Availability Checker**: A dedicated tool to verify floor availability (Ground/Second) for specific date ranges before booking.
*   **Flexible Registration**: 
    *   Custom financial matrices for personal bookings (Base price + Extra guest fees).
    *   Automatic calculation of total bills, advances, and balance dues.
*   **WhatsApp Integration**: Direct buttons to dispatch "Advance Payment Requests" and "Checkout Bills" to guests.
*   **Ratings & Notes**: Track guest satisfaction and internal operational comments.

### 🛡️ Secure QID Registry
*   **Identity Management**: A secure repository for verified guest profiles and ID identifiers.
*   **Storage Monitor**: Tracks Firestore and media storage usage against the 5GB free tier, including an "Optimize Storage" utility.
*   **Batch Operations**: Supports batch deletion and comprehensive profile modification.
*   **Secure Previews**: In-app lightbox for viewing guest ID assets safely.

### ⚙️ Automation & Synchronization
*   **Airbnb Sync**: One-click sync to update the dashboard with the latest Airbnb email confirmations and bookings.
*   **CI/CD Pipeline**: Fully automated deployment via GitHub Actions with integrated **Short.io** URL alignment.

## 🛠️ Tech Stack
*   **Frontend**: HTML5, CSS3, Bootstrap 5.3, Bootstrap Icons.
*   **Backend**: Google Apps Script (JavaScript).
*   **Deployment**: GitHub Actions, Clasp CLI.

## 📝 CI/CD Workflow Details
The project uses a two-stage deployment pipeline:
1.  **Test Stage**: Pushes code to the `@HEAD` deployment for immediate testing in a development environment.
2.  **Production Deploy**: Triggered after the test stage, it updates the live production deployment ID with a new version snapshot and provides a summarized deployment report.

## 🛤️ Pending Features & Roadmap
*   **Data Visualization**: Integrated charts to visualize monthly revenue growth and occupancy trends.
*   **Automated Emailers**: Triggering automated confirmation emails to guests upon registration.

## 📦 Getting Started

### Prerequisites
*   Node.js (v20+)
*   `@google/clasp` installed globally (`npm install -g @google/clasp`)

### GitHub Secrets Setup
To enable the automated deployment, ensure the following are set in your repository secrets:
*   `CLASPRC_JSON`: Your Google account authentication credentials.
*   `SCRIPT_ID`: The unique ID of your Google Apps Script project.
*   `SHORT_IO_API_KEY`: API key from Short.io for vanity URL updates.

### GitHub Environment Variables
Set these in Settings > Environments:
*   `HEAD_DEPLOYMENT_ID`: For dev testing.
*   `DEPLOYMENT_ID`: For the live production app.

---
*Maintained by the Vinyasa Nilaya Engineering Team.*