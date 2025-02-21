import { ReactNode } from 'react';

interface RootLayoutProps {
    children: ReactNode; // Define the type of the `children` prop
}

export default function RootLayout({ children }: RootLayoutProps) {
    return (
        <html lang="en">
            <body>{children}</body>
        </html>
    );
}