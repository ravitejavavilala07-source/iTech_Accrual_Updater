import React from "react";

type Variant = "heroSecondary" | "ghost";

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: Variant;
}

const variantClasses: Record<Variant, string> = {
  heroSecondary:
    "liquid-glass text-foreground/95 hover:text-foreground transition-colors font-medium",
  ghost:
    "text-foreground/90 hover:text-foreground transition-colors bg-transparent",
};

export function Button({
  variant = "ghost",
  className = "",
  children,
  ...props
}: ButtonProps) {
  return (
    <button
      className={`${variantClasses[variant]} ${className}`}
      {...props}
    >
      {children}
    </button>
  );
}
