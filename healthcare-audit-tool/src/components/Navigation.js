import React from 'react';
import { Link } from 'react-router-dom';

function Navigation() {
    return (
        <nav>
            <ul>
                <li><Link to="/odag">ODAG</Link></li>
                <li><Link to="/cdag">CDAG</Link></li>
                <li><Link to="/snpcc">SNPCC</Link></li>
                <li><Link to="/cpe">CPE</Link></li>
                <li><Link to="/fa">FA</Link></li>
            </ul>
        </nav>
    );
}

export default Navigation;
