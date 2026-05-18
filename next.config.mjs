/** @type {import('next').NextConfig} */
const nextConfig = {
  // pdfjs-dist resolves its worker relative to its own package files; if Next
  // bundles it into the route, that resolution breaks. Keep it external.
  serverExternalPackages: ['pdfjs-dist'],
};

export default nextConfig;
