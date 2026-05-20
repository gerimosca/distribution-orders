/** @type {import('next').NextConfig} */
const nextConfig = {
  // pdfjs-dist resolves its worker relative to its own package files; if Next
  // bundles it into the route, that resolution breaks. Keep it external.
  serverExternalPackages: ['pdfjs-dist'],
  // The worker is loaded dynamically (import(pdf.worker.mjs)), so Next's
  // file tracer doesn't see it and Vercel ships the function without it
  // ("Cannot find module .../pdf.worker.mjs"). Force-include it.
  outputFileTracingIncludes: {
    '/api/**/*': ['./node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs'],
  },
};

export default nextConfig;
