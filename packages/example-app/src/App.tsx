import React from 'react';
import logo from './logo.svg';
import './App.css';
import { Outlet, Route, RouterProvider, createBrowserRouter, createRoutesFromElements } from 'react-router-dom';
import { NotFoundPage } from './page/notFoundPage';
import { Example1Page } from './page/example1Page';
import { Example2Page } from './page/example2Page';
import { Example3Page } from './page/example3Page';

function App() {
  return (
    <RouterProvider router={createBrowserRouter(createRoutesFromElements(
      <Route 
        path='/'
        errorElement={<NotFoundPage />}
        element={<Outlet />}
      >
        <Route path="/example1" element={<Example1Page />} />
        <Route path="/example2" element={<Example2Page />} />
        <Route path="/example3" element={<Example3Page />} />
      </Route>
    ), {basename: "/ts-exceljs-utility"})} />
  );
}

export default App;
