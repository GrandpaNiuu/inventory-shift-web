/** @type {import('next').NextConfig} */
const isGitHubPages = process.env.GITHUB_ACTIONS === "true";

const nextConfig = {
  reactStrictMode: true,
  output: "export",
  basePath: isGitHubPages ? "/inventory-shift-web" : "",
  assetPrefix: isGitHubPages ? "/inventory-shift-web/" : "",
  images: {
    unoptimized: true
  }
};

export default nextConfig;
