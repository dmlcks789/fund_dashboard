import { useEffect, useRef, useState } from "react";

export default function OverflowTooltipText({ value }: { value: string }) {
  const textRef = useRef<HTMLSpanElement>(null);
  const [isOverflow, setIsOverflow] = useState(false);

  useEffect(() => {
    const element = textRef.current;
    if (!element) {
      return;
    }

    const checkOverflow = () => {
      setIsOverflow(element.scrollWidth > element.clientWidth + 1);
    };

    checkOverflow();
    window.addEventListener("resize", checkOverflow);

    let observer: ResizeObserver | null = null;
    if (typeof ResizeObserver !== "undefined") {
      observer = new ResizeObserver(checkOverflow);
      observer.observe(element);
    }

    return () => {
      window.removeEventListener("resize", checkOverflow);
      observer?.disconnect();
    };
  }, [value]);

  return (
    <span ref={textRef} className="cellEllipsisText" title={isOverflow ? value : undefined}>
      {value}
    </span>
  );
}
