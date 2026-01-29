# The Market Masters (TMM) — Static Website

This repository contains the official **multi-page static website** for **The Market Masters (TMM)**, a branding and marketing agency based in **Maiduguri, Borno State, Nigeria**. The site is intentionally built as a lightweight, framework-free project to ensure speed, clarity, and ease of maintenance on standard shared hosting (cPanel-friendly).

The website uses **HTML5**, **CSS3**, **Vanilla JavaScript**, **Bootstrap 5 (CDN)**, **Bootstrap Icons**, and **Google Fonts (Fraunces for headings, Inter for body text)**. There is no backend, database, or build process required.

---

## Project Structure

tmm-site/
├── index.html  
├── about.html  
├── services.html  
├── events.html  
├── contact.html  
├── community.html  
├── README.md  
└── assets/  
  ├── css/  
  │   └── styles.css  
  ├── js/  
  │   └── main.js  
  └── img/  
    └── logo.jpg  

---

## Pages

- **Home (`index.html`)**: Hero section with rotating headline text, services preview, events preview, and primary CTAs.
- **About (`about.html`)**: Overview of TMM, mission, values, and positioning.
- **Services (`services.html`)**: Detailed service offerings presented in a clean, modern layout.
- **Events (`events.html`)**: Full list of events and activations with automatic filtering by category.
- **Contact (`contact.html`)**: Contact information cards (address, phone, email), social links, and Google Maps embed.
- **Community (`community.html`)**: Call-to-action page linking to a Google Form for community sign-ups.

---

## Local Development

The site can be opened directly in a browser, but running a local server is recommended.

Using VS Code Live Server:
- Install the **Live Server** extension.
- Right-click `index.html` and select **Open with Live Server**.

---

## Deployment

Upload the contents of the tmm-site/ directory to your hosting root (usually public_html/). Ensure index.html is directly inside the root directory. Once uploaded, the site will be accessible via the domain (e.g. https://themarketmasters.com.ng).

---

## Configuration

# Contact Information

Edit phone numbers, email address, physical address, and Google Maps embed URL directly in contact.html.

# Social Media Links

Update all social media URLs in the footer and contact sections to reflect the correct handles.

# Events Management

All events are managed from a single data source to keep the Home page and Events page in sync.

Events are defined in assets/js/main.js inside the eventsData array. Each event object follows this structure:

{
  title: "Event name",
  date: "Schedule",
  location: "Location",
  tag: "Category",
  desc: "Short description"
}

To add a new event, insert a new object into the eventsData array. Events automatically:

Appear on the Home page preview

Appear on the Events page

Generate a new filter button if a new category (tag) is introduced

Recommended tag values: Conference, Experiential, Community, Marketplace, Digital, Webinar, Workshop.

---

### Dynamic Hero Text

The Home page hero headline includes rotating text controlled via data attributes in index.html. The phrases and rotation speed can be changed by editing the data-rotate and data-interval values on the hero text element.

# Styling & Branding

Global styles are defined in assets/css/styles.css. Brand colors, spacing, typography, and reusable components are centralized using CSS variables under the :root selector. Updating colors or typography in this file will update the entire site.

The site logo is located at assets/img/logo.jpg. Replace this file to update the logo without modifying HTML.

# Accessibility & UX

The site uses semantic HTML, accessible navigation, keyboard-friendly interactions, skip-to-content support, and subtle motion effects for a polished and inclusive user experience.

---

### Maintenance Notes

This is a static website with no backend dependencies. Any developer familiar with HTML, CSS, and JavaScript can maintain or extend the site. The structure is intentionally simple to support long-term scalability without technical overhead.