import './globals.css';

export const metadata = {
  title: 'Visual Data AI Assistant',
  description: 'Low-cost structured extraction for documents, fields, and tables.',
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
