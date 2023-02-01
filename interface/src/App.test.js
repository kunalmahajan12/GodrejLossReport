import { render, screen } from '@testing-library/react';
import App from './App';

test('renders learn react link', () => {
  render(<App />);
  const titleElement = screen.getByText(/LOSS CAPTURING SYSTEM/i);
  expect(titleElement).toBeInTheDocument();
});
