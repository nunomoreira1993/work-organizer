import type { MetadataRoute } from "next";

export default function manifest(): MetadataRoute.Manifest {
  return {
    name: "Work Organizer",
    short_name: "Organizer",
    description: "Planeamento e registo de trabalho integrado com GitLab e Outlook.",
    start_url: "/",
    display: "standalone",
    background_color: "#f4f7fb",
    theme_color: "#0c79d8",
    icons: [
      {
        src: "/favicon.svg",
        sizes: "any",
        type: "image/svg+xml",
      },
    ],
  };
}
