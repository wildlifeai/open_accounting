/**
 * TrackingBuilder.js
 * Merges the three layers for one funding source's quarterly tracking screen:
 *   baseline (frozen budget)  + actual (Xero)  + forecast (GM-editable).
 *
 * Per milestone x quarter the "effective" expected spend is:
 *   - past quarter      -> actual
 *   - current / future  -> forecast override if the GM set one, else baseline
 * Variance = expected total - baseline total (positive = heading over budget).
 *
 * Baseline + actual come from the cached snapshot (compact, from buildTracking_);
 * the forecast is read live so the GM's edits appear immediately.
 */

function composeTrackingForSource(trackingSource, forecastMap, currentQuarter) {
  const quarters = trackingSource.quarters.slice();
  const curNum = quarterSortNum(currentQuarter);

  const milestones = trackingSource.milestones.map(m => {
    const cells = quarters.map(q => {
      const baseline = m.baseline[q] || 0;
      const actual = m.actual[q] || 0;
      const fkey = trackingSource.source + '||' + m.item + '||' + q;
      const override = forecastMap[fkey] ? forecastMap[fkey].cost : null;
      const isPast = quarterSortNum(q) < curNum;
      const forecast = override !== null ? override : baseline;
      const effective = isPast ? actual : forecast;
      return { quarter: q, baseline: baseline, actual: actual,
        forecast: forecast, hasOverride: override !== null,
        effective: Math.round(effective), editable: !isPast };
    });
    const baselineTotal = sum_(cells, 'baseline');
    const expectedTotal = sum_(cells, 'effective');
    return {
      item: m.item, milestone: m.milestone, cells: cells,
      baselineTotal: Math.round(baselineTotal),
      actualToDate: Math.round(sum_(cells, 'actual')),
      expectedTotal: Math.round(expectedTotal),
      variance: Math.round(expectedTotal - baselineTotal)
    };
  });

  // Column + grand totals.
  const colTotals = quarters.map((q, i) => ({
    quarter: q,
    baseline: Math.round(milestones.reduce((t, m) => t + m.cells[i].baseline, 0)),
    actual: Math.round(milestones.reduce((t, m) => t + m.cells[i].actual, 0)),
    effective: Math.round(milestones.reduce((t, m) => t + m.cells[i].effective, 0))
  }));

  return {
    source: trackingSource.source, status: trackingSource.status,
    project: trackingSource.project, currentQuarter: currentQuarter,
    quarters: quarters, milestones: milestones, colTotals: colTotals,
    totals: {
      baseline: Math.round(milestones.reduce((t, m) => t + m.baselineTotal, 0)),
      actual: Math.round(milestones.reduce((t, m) => t + m.actualToDate, 0)),
      expected: Math.round(milestones.reduce((t, m) => t + m.expectedTotal, 0)),
      variance: Math.round(milestones.reduce((t, m) => t + m.variance, 0))
    }
  };
}

function sum_(cells, field) {
  return cells.reduce((t, c) => t + (c[field] || 0), 0);
}
