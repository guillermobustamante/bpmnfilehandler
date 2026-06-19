import type { ButtonHTMLAttributes, ReactNode } from "react";

type IconButtonProps = ButtonHTMLAttributes<HTMLButtonElement> & {
  children: ReactNode;
  label: string;
};

export function IconButton({ children, className = "", label, ...buttonProps }: IconButtonProps) {
  return (
    <button
      {...buttonProps}
      aria-label={label}
      className={`icon-button ${className}`.trim()}
      title={label}
      type={buttonProps.type || "button"}
    >
      {children}
    </button>
  );
}

