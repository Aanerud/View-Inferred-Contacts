# Outlook Add-In: View Inferred Contacts

## Overview

This Outlook add-in enables users to view inferred contact information from email messages using Exchange Web Services (EWS). The add-in fetches and displays data like task existence and detailed contact information in vCard format.

## Features

- **EntityExtraction/HasTask:** Determines if the email has associated tasks.
- **EntityExtraction/ContactInformation:** Extracts and formats contact information into vCard format.
  
## Project Structure

- **taskpane.css:** Stylesheet for the task pane UI.
- **taskpane.html:** HTML structure for the task pane.
- **taskpane.js:** JavaScript file that handles the EWS requests and data extraction.
- **manifest.xml:** Manifest file to configure the Outlook add-in.
- **webpack.config.js:** Webpack configuration for building the project.
- **package.json:** Node.js dependencies and scripts.

## Getting Started

### Prerequisites

- Node.js (version 14 or above)
- Outlook 2016 or later
- Exchange Web Services (EWS) enabled on your Exchange server

### Setup Instructions

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/aanerud/outlook-inferred-contacts.git
   cd outlook-inferred-contacts
   ```

2. **Install Dependencies:**

   ```bash
   npm install
   ```

3. **Run the Add-In Locally:**

   Use the following command to start the development server and sideload the add-in:

   ```bash
   npm run start
   ```

   This command will open Outlook, and the add-in will be available in the message read and compose surfaces.

4. **Build for Production:**

   When you're ready to deploy, run:

   ```bash
   npm run build
   ```

### Usage

1. **Sideload the Add-In:**

   Open Outlook and ensure the add-in is sideloaded. Follow the prompts from the `npm start` command or manually sideload the `manifest.xml` file.

2. **Get Extracted Data:**

   - Open any email message.
   - Click on the "Get extracted data" button.
   - The task pane will display whether the email has associated tasks and will extract any available contact information, displaying it in vCard format.

### Technical Details

- **EWS Requests:** The add-in sends two separate EWS requests to fetch the `HasTask` and `ContactInformation` properties.
- **Parsing and Display:** The contact information is parsed into a JSON object, then formatted into vCard format and displayed in the task pane.

### Development

To modify and extend the add-in:

1. Edit the `taskpane.js` file to change the logic for extracting and displaying data.
2. Use `taskpane.css` to update the styling of the add-in UI.
3. Make sure to update `webpack.config.js` if you add new assets or change the structure of the add-in.

### Troubleshooting

- **Common Errors:**
  - "Error retrieving the HasTask/ContactInformation property": Ensure your EWS service is correctly configured and accessible.
  - "Error parsing contact information": This typically occurs due to unexpected data structures. Check the EWS response in the logs.

- **Logs:** All requests and responses are logged in the console for debugging purposes.

### License

This project is licensed under the MIT License.

## Contributions

Contributions are welcome! Please fork this repository and submit a pull request with your changes.

## Contact

For questions, please reach out to [aaanerud@microsoft.com](mailto:aaanerud@microsoft.com).
