import { Navbar } from "./components/Navbar";
import { AccrualForm } from "./components/AccrualForm";
import { ToastHost } from "./components/Toast";

const VIDEO_URL =
  "https://d8j0ntlcm91z4.cloudfront.net/user_38xzZboKViGWJOttwIXH07lWA1P/hf_20260328_065045_c44942da-53c6-4804-b734-f9e07fc22e08.mp4";

export default function App() {
  return (
    <div className="relative min-h-screen overflow-hidden bg-background">
      {/* Background video (muted, auto-loop via video loop attribute) */}
      <video
        src={VIDEO_URL}
        className="absolute inset-0 w-full h-full object-cover opacity-25"
        autoPlay
        muted
        loop
        playsInline
      />

      {/* Dark overlay for readability */}
      <div className="absolute inset-0 bg-background/80 pointer-events-none" />

      <div className="relative z-10 min-h-screen flex flex-col">
        <Navbar />
        <main className="flex-1">
          <AccrualForm />
        </main>
      </div>
      <ToastHost />
    </div>
  );
}
