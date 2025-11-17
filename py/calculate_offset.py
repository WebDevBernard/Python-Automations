# Example usage:
pattern = (
    217.20001220703125,
    54.4100227355957,
    548.9254150390625,
    64.04902648925781,
)
target = (352, 67.12498474121094, 467.5700378417969, 79.4909896850586)


def offset_below(pattern_rect, target_rect):
    """
    Calculate offset for target rect below pattern rect.
    Prints offset values to copy into constants.

    Args:
        pattern_rect: tuple (x0, y0, x1, y1) - where pattern is found
        target_rect: tuple (x0, y0, x1, y1) - where you want to extract
    """
    p_x0, p_y0, p_x1, p_y1 = pattern_rect
    t_x0, t_y0, t_x1, t_y1 = target_rect

    dx0 = t_x0 - p_x0
    dy0 = t_y0 - p_y0  # distance from pattern's BOTTOM edge
    dx1 = t_x1 - p_x1
    dy1 = t_y1 - p_y1  # distance from pattern's BOTTOM edge

    print(f"dx0={dx0:.2f}, dy0={dy0:.2f}, dx1={dx1:.2f}, dy1={dy1:.2f}")


def offset_right(pattern_rect, target_rect):
    """
    Calculate offset for target rect to the right of pattern rect.
    Prints offset values to copy into constants.
    Args:
        pattern_rect: tuple (x0, y0, x1, y1) - where pattern is found
        target_rect: tuple (x0, y0, x1, y1) - where you want to extract
    """
    p_x0, p_y0, p_x1, p_y1 = pattern_rect
    t_x0, t_y0, t_x1, t_y1 = target_rect

    # Calculate from pattern's LEFT edge (x0) to work with extract_with_pattern_and_offset
    dx0 = t_x0 - p_x0  # distance from pattern's LEFT edge to target's left
    dy0 = t_y0 - p_y0
    dx1 = t_x1 - p_x1  # keeps the width calculation correct
    dy1 = t_y1 - p_y1

    print(f"dx0={dx0:.2f}, dy0={dy0:.2f}, dx1={dx1:.2f}, dy1={dy1:.2f}")


def offset_above(pattern_rect, target_rect):
    """
    Calculate offset for target rect above pattern rect.
    Prints offset values to copy into constants.
    Args:
        pattern_rect: tuple (x0, y0, x1, y1) - where pattern is found
        target_rect: tuple (x0, y0, x1, y1) - where you want to extract
    """
    p_x0, p_y0, p_x1, p_y1 = pattern_rect
    t_x0, t_y0, t_x1, t_y1 = target_rect

    dx0 = t_x0 - p_x0
    dy0 = t_y0 - p_y0  # distance from pattern's TOP edge
    dx1 = t_x1 - p_x1
    dy1 = t_y1 - p_y0  # distance from pattern's TOP edge

    print(f"dx0={dx0:.2f}, dy0={dy0:.2f}, dx1={dx1:.2f}, dy1={dy1:.2f}")


print("Above:")
offset_above(pattern, target)

print("\nBelow:")
offset_below(pattern, target)

print("\nRight:")
offset_right(pattern, target)
