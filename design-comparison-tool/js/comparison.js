export function highlightMatches(containers, term) {
  let count = 0;

  for (const container of containers) {
    let first = null;

    for (const cell of container.querySelectorAll('td')) {
      const match =
        Boolean(term) &&
        cell.textContent
          .toLocaleLowerCase()
          .includes(term.toLocaleLowerCase());

      cell.classList.toggle('match', match);

      if (match) {
        count++;
        if (!first) first = cell;
      }
    }

    if (first) {
      first.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
        inline: 'nearest',
      });
    }
  }

  return count;
}
