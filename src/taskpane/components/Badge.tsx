interface BadgeProps {
  count: number;
}

export function Badge({ count }: BadgeProps) {
  if (count === 0) return null;
  return (
    <span
      class="dp-badge"
      aria-label={`${count} issue${count === 1 ? '' : 's'}`}
    >
      {count > 99 ? '99+' : count}
    </span>
  );
}
